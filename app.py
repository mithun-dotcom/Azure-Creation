"""
Shared Mailbox Provisioner — backend.

Single-file FastAPI app that:
  - Accepts Microsoft 365 app-registration credentials (client ID/secret + .pfx cert)
  - Lists licensed users via Microsoft Graph
  - Bulk-creates shared mailboxes via Exchange Online PowerShell
  - Grants Full Access + Send-on-Behalf to a chosen licensed user
  - Sets password and unblocks sign-in via PowerShell (Set-MsolUserPassword / MSOnline)
  - Flips the org-wide "Turn off SMTP AUTH" checkbox to unchecked (enables SMTP AUTH)

Deploy on Render using the supplied Dockerfile (which installs pwsh and the
ExchangeOnlineManagement module). Sessions and jobs are kept in memory.
"""

import asyncio
import csv
import io
import json
import os
import secrets
import subprocess
import tempfile
from contextlib import asynccontextmanager
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional

import httpx
from fastapi import (
    BackgroundTasks,
    FastAPI,
    File,
    Form,
    HTTPException,
    UploadFile,
)
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response
from pydantic import BaseModel

# ---------------------------------------------------------------------------
# In-memory state
# ---------------------------------------------------------------------------

SESSIONS: Dict[str, dict] = {}
JOBS: Dict[str, dict] = {}
SESSION_TTL = timedelta(hours=2)
GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def _purge_session(token: str) -> None:
    sess = SESSIONS.pop(token, None)
    if sess and sess.get("cert_path") and os.path.exists(sess["cert_path"]):
        try:
            os.remove(sess["cert_path"])
        except OSError:
            pass


async def _cleanup_loop() -> None:
    while True:
        try:
            await asyncio.sleep(300)
            now = datetime.utcnow()
            for token in [t for t, s in SESSIONS.items() if s["expires_at"] < now]:
                _purge_session(token)
        except asyncio.CancelledError:
            break
        except Exception:
            await asyncio.sleep(5)


@asynccontextmanager
async def lifespan(app: FastAPI):
    task = asyncio.create_task(_cleanup_loop())
    try:
        yield
    finally:
        task.cancel()


app = FastAPI(title="Shared Mailbox Provisioner", lifespan=lifespan)

_allowed = os.environ.get("ALLOWED_ORIGINS", "*").split(",")
app.add_middleware(
    CORSMiddleware,
    allow_origins=_allowed,
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ---------------------------------------------------------------------------
# Microsoft Graph helpers
# ---------------------------------------------------------------------------


async def get_graph_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
    }
    async with httpx.AsyncClient(timeout=30) as client:
        r = await client.post(url, data=data)
    if r.status_code != 200:
        raise HTTPException(401, f"Graph auth failed: {r.text}")
    return r.json()["access_token"]


async def get_tenant_domain(token: str) -> str:
    """
    Exchange Online's -Organization parameter requires the tenant domain name
    (e.g. contoso.onmicrosoft.com), NOT the directory GUID.
    """
    async with httpx.AsyncClient(timeout=30) as client:
        r = await client.get(
            f"{GRAPH_BASE}/organization",
            headers={"Authorization": f"Bearer {token}"},
            params={"$select": "verifiedDomains"},
        )
    if r.status_code >= 400:
        raise HTTPException(r.status_code, f"Graph /organization error: {r.text}")
    domains = r.json()["value"][0]["verifiedDomains"]
    for d in domains:
        if d["name"].endswith(".onmicrosoft.com") and d.get("isDefault"):
            return d["name"]
    for d in domains:
        if d["name"].endswith(".onmicrosoft.com"):
            return d["name"]
    raise HTTPException(400, "Could not find .onmicrosoft.com domain for this tenant")


async def list_licensed_users(token: str) -> List[dict]:
    """Return all directory users that have at least one assigned license."""
    users: List[dict] = []
    params = {
        "$select": "id,displayName,userPrincipalName,mail,assignedLicenses,accountEnabled",
        "$top": "999",
    }
    url: Optional[str] = f"{GRAPH_BASE}/users"
    async with httpx.AsyncClient(timeout=60) as client:
        first = True
        while url:
            r = await client.get(
                url,
                headers={"Authorization": f"Bearer {token}"},
                params=params if first else None,
            )
            if r.status_code >= 400:
                raise HTTPException(r.status_code, f"Graph users error: {r.text}")
            data = r.json()
            for u in data.get("value", []):
                if u.get("assignedLicenses"):
                    users.append(
                        {
                            "id": u["id"],
                            "displayName": u.get("displayName") or u["userPrincipalName"],
                            "userPrincipalName": u["userPrincipalName"],
                            "mail": u.get("mail"),
                            "accountEnabled": u.get("accountEnabled", False),
                        }
                    )
            url = data.get("@odata.nextLink")
            first = False
    return users


# ---------------------------------------------------------------------------
# PowerShell helpers
# ---------------------------------------------------------------------------


def ps_escape(s: str) -> str:
    """Escape for a PowerShell single-quoted string literal."""
    if s is None:
        return ""
    return str(s).replace("'", "''")


def _connect_block(cert_path: str, cert_password: str, organization: str, app_id: str) -> str:
    """
    Connect to Exchange Online using certificate auth.
    -Organization must be the tenant domain name, NOT the GUID.
    """
    return f"""
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$securePass = ConvertTo-SecureString -String '{ps_escape(cert_password)}' -AsPlainText -Force
Import-Module ExchangeOnlineManagement -ErrorAction Stop | Out-Null
Connect-ExchangeOnline -CertificateFilePath '{cert_path}' -CertificatePassword $securePass -AppId '{ps_escape(app_id)}' -Organization '{ps_escape(organization)}' -ShowBanner:$false -ErrorAction Stop | Out-Null
"""


def run_pwsh(script: str, timeout: int = 120) -> tuple[int, str, str]:
    proc = subprocess.run(
        ["pwsh", "-NoProfile", "-NonInteractive", "-Command", script],
        capture_output=True,
        text=True,
        timeout=timeout,
    )
    return proc.returncode, proc.stdout, proc.stderr


# ---------------------------------------------------------------------------
# Models
# ---------------------------------------------------------------------------


class DistributionEntry(BaseModel):
    licensedUserUpn: str
    count: int


class JobRequest(BaseModel):
    mailboxes: List[Dict[str, Any]]
    distribution: List[DistributionEntry]


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------


@app.get("/")
def root():
    return {"name": "Shared Mailbox Provisioner", "status": "ok"}


@app.get("/api/health")
def health():
    try:
        rc, out, err = run_pwsh("$PSVersionTable.PSVersion.ToString()", timeout=10)
        pwsh_version = out.strip() if rc == 0 else None
    except Exception as e:
        pwsh_version = f"unavailable: {e}"
    return {
        "status": "ok",
        "pwsh": pwsh_version,
        "sessions": len(SESSIONS),
        "jobs": len(JOBS),
    }


@app.post("/api/session")
async def create_session(
    tenantId: str = Form(...),
    clientId: str = Form(...),
    clientSecret: str = Form(...),
    certPassword: str = Form(...),
    certFile: UploadFile = File(...),
):
    """
    Validate credentials end-to-end:
      1. Acquire Graph token
      2. Resolve tenant GUID → .onmicrosoft.com domain
      3. Test Exchange Online connection with cert
    """
    cert_bytes = await certFile.read()
    if not cert_bytes:
        raise HTTPException(400, "Certificate file is empty")
    fd, cert_path = tempfile.mkstemp(suffix=".pfx", prefix="m365_")
    with os.fdopen(fd, "wb") as f:
        f.write(cert_bytes)

    try:
        gt = await get_graph_token(tenantId, clientId, clientSecret)
    except HTTPException:
        os.remove(cert_path)
        raise

    try:
        organization = await get_tenant_domain(gt)
    except HTTPException:
        os.remove(cert_path)
        raise

    test_script = (
        _connect_block(cert_path, certPassword, organization, clientId)
        + "Get-OrganizationConfig | Select-Object -First 1 -Property Name | "
          "ConvertTo-Json -Compress\n"
          "Disconnect-ExchangeOnline -Confirm:$false | Out-Null\n"
    )
    rc, out, err = run_pwsh(test_script, timeout=180)
    if rc != 0:
        os.remove(cert_path)
        msg = (err or out).strip()
        if len(msg) > 1500:
            msg = msg[:1500] + "…"
        raise HTTPException(401, f"Exchange Online connect failed: {msg}")

    token = secrets.token_urlsafe(32)
    SESSIONS[token] = {
        "tenant_id": tenantId,
        "organization": organization,
        "client_id": clientId,
        "client_secret": clientSecret,
        "cert_password": certPassword,
        "cert_path": cert_path,
        "expires_at": datetime.utcnow() + SESSION_TTL,
    }
    return {
        "sessionToken": token,
        "expiresAt": SESSIONS[token]["expires_at"].isoformat() + "Z",
    }


def get_session(token: str) -> dict:
    sess = SESSIONS.get(token)
    if not sess:
        raise HTTPException(401, "Session not found")
    if sess["expires_at"] < datetime.utcnow():
        _purge_session(token)
        raise HTTPException(401, "Session expired")
    return sess


@app.delete("/api/session/{token}")
def delete_session(token: str):
    _purge_session(token)
    return {"ok": True}


@app.get("/api/session/{token}/users/licensed")
async def get_licensed_users(token: str):
    sess = get_session(token)
    gt = await get_graph_token(sess["tenant_id"], sess["client_id"], sess["client_secret"])
    return {"users": await list_licensed_users(gt)}


@app.post("/api/session/{token}/job")
async def start_job(token: str, body: JobRequest, background_tasks: BackgroundTasks):
    sess = get_session(token)
    total_dist = sum(e.count for e in body.distribution)
    if total_dist != len(body.mailboxes):
        raise HTTPException(
            400,
            f"Distribution sums to {total_dist} but you uploaded {len(body.mailboxes)} mailboxes.",
        )
    if not body.mailboxes:
        raise HTTPException(400, "No mailboxes to create.")

    assignments: List[dict] = []
    idx = 0
    for entry in body.distribution:
        for _ in range(entry.count):
            assignments.append({"mailbox": body.mailboxes[idx], "owner": entry.licensedUserUpn})
            idx += 1

    job_id = secrets.token_urlsafe(16)
    JOBS[job_id] = {
        "id": job_id,
        "status": "queued",
        "control": None,
        "logs": [],
        "results": [],
        "total": len(assignments),
        "completed": 0,
        "started_at": datetime.utcnow().isoformat() + "Z",
        "finished_at": None,
        "session_token": token,
    }
    background_tasks.add_task(_run_job, job_id, assignments)
    return {"jobId": job_id}


@app.get("/api/session/{token}/job/{job_id}")
def get_job(token: str, job_id: str):
    get_session(token)
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(404, "Job not found")
    return job


@app.get("/api/session/{token}/job/{job_id}/report")
def get_job_report(token: str, job_id: str, includePassword: int = 0):
    get_session(token)
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(404, "Job not found")
    out = io.StringIO()
    w = csv.writer(out)
    header = [
        "Index",
        "PrimarySmtpAddress",
        "AssignedOwner",
        "Created",
        "FullAccess",
        "SendAs",
        "SendOnBehalf",
        "SignInUnblocked",
        "PasswordSet",
        "Error",
    ]
    if includePassword:
        header.append("Password")
    w.writerow(header)
    for r in job.get("results", []):
        row = [
            r.get("idx"),
            r.get("smtp", ""),
            r.get("owner", ""),
            r.get("create") or "",
            r.get("fullaccess") or "",
            r.get("sendas") or "",
            r.get("sendonbehalf") or "",
            r.get("signinUnblocked") or "",
            r.get("passwordSet") or "",
            r.get("error") or "",
        ]
        if includePassword:
            row.append(r.get("password") or "")
        w.writerow(row)
    suffix = "-with-passwords" if includePassword else ""
    return Response(
        content=out.getvalue(),
        media_type="text/csv",
        headers={"Content-Disposition": f"attachment; filename=mailbox-job-{job_id}{suffix}.csv"},
    )


@app.post("/api/session/{token}/job/{job_id}/pause")
def pause_job(token: str, job_id: str):
    get_session(token)
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(404, "Job not found")
    if job["status"] not in ("queued", "running"):
        raise HTTPException(400, f"Cannot pause job in status '{job['status']}'")
    job["control"] = "pause"
    return {"status": "pause-requested"}


@app.post("/api/session/{token}/job/{job_id}/resume")
def resume_job(token: str, job_id: str):
    get_session(token)
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(404, "Job not found")
    if job["status"] != "paused":
        raise HTTPException(400, f"Cannot resume job in status '{job['status']}'")
    job["control"] = "resume"
    job["status"] = "running"
    return {"status": "resumed"}


@app.post("/api/session/{token}/job/{job_id}/stop")
def stop_job(token: str, job_id: str):
    get_session(token)
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(404, "Job not found")
    if job["status"] not in ("queued", "running", "paused"):
        raise HTTPException(400, f"Cannot stop job in status '{job['status']}'")
    job["control"] = "stop"
    return {"status": "stop-requested"}


# ---------------------------------------------------------------------------
# Job runner
# ---------------------------------------------------------------------------


def _log(job_id: str, level: str, message: str, **extra) -> None:
    JOBS[job_id]["logs"].append(
        {
            "ts": datetime.utcnow().isoformat() + "Z",
            "level": level,
            "message": message,
            **extra,
        }
    )


def _normalize_row(row: Dict[str, Any]) -> Dict[str, str]:
    """Allow common alternative column names from the CSV."""
    pick = lambda *names: next(
        (str(row[n]) for n in names if n in row and row[n] not in (None, "")), ""
    )
    smtp = pick("PrimarySmtpAddress", "Email", "EmailAddress", "UserPrincipalName", "UPN")
    return {
        "DisplayName": pick("DisplayName", "Display Name", "Name") or smtp.split("@")[0],
        "Alias": pick("Alias", "MailNickname") or smtp.split("@")[0],
        "PrimarySmtpAddress": smtp,
        "Password": pick("Password", "Pass"),
    }


async def _run_job(job_id: str, assignments: List[dict]) -> None:
    job = JOBS[job_id]
    sess = SESSIONS.get(job["session_token"])
    if not sess:
        job["status"] = "failed"
        _log(job_id, "error", "Session expired before job started")
        job["finished_at"] = datetime.utcnow().isoformat() + "Z"
        return

    job["status"] = "running"
    _log(job_id, "info", f"Job started — {len(assignments)} mailboxes to provision")

    cert_path = sess["cert_path"]
    cert_pw = sess["cert_password"]
    tenant = sess["tenant_id"]
    organization = sess["organization"]
    app_id = sess["client_id"]

    # ---------------------------------------------------------------------------
    # Build PowerShell script:
    # Phase 1 — create mailbox, set permissions (Exchange Online)
    # Phase 2 — set password + enable sign-in via Set-AzureADUser (MSOnline/AzureAD)
    #            falling back to Update-MgUser (Microsoft.Graph) if available
    # Phase 3 — enable org-wide SMTP AUTH
    # All in ONE PowerShell session to avoid reconnection overhead.
    # ---------------------------------------------------------------------------

    def build_action(i: int, a: dict) -> str:
        m = _normalize_row(a["mailbox"])
        owner = a["owner"]
        smtp = m["PrimarySmtpAddress"]
        password = m["Password"]
        return f"""
$idx = {i}
$smtp = '{ps_escape(smtp)}'
$owner = '{ps_escape(owner)}'
$pw = '{ps_escape(password)}'
try {{
    # --- Create shared mailbox ---
    $mbx = New-Mailbox -Shared -Name '{ps_escape(m["DisplayName"])}' `
        -DisplayName '{ps_escape(m["DisplayName"])}' `
        -Alias '{ps_escape(m["Alias"])}' `
        -PrimarySmtpAddress $smtp -ErrorAction Stop
    Write-Host (ConvertTo-Json -Compress @{{ stage='create'; idx=$idx; smtp=$smtp; ok=$true; upn=$mbx.UserPrincipalName }})

    # --- Full Access ---
    Add-MailboxPermission -Identity $smtp -User $owner `
        -AccessRights FullAccess -InheritanceType All `
        -AutoMapping $true -ErrorAction Stop | Out-Null
    Write-Host (ConvertTo-Json -Compress @{{ stage='fullaccess'; idx=$idx; smtp=$smtp; ok=$true }})

    # --- Send As ---
    Add-RecipientPermission -Identity $smtp -Trustee $owner `
        -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
    Write-Host (ConvertTo-Json -Compress @{{ stage='sendas'; idx=$idx; smtp=$smtp; ok=$true }})

    # --- Send on Behalf ---
    Set-Mailbox -Identity $smtp -GrantSendOnBehalfTo @{{ Add=$owner }} -ErrorAction Stop
    Write-Host (ConvertTo-Json -Compress @{{ stage='sendonbehalf'; idx=$idx; smtp=$smtp; ok=$true }})

    # --- Password + enable sign-in via PowerShell ---
    # Wait briefly for the account to fully propagate
    Start-Sleep -Seconds 5
    if ($pw -ne '') {{
        $securePw = ConvertTo-SecureString -String $pw -AsPlainText -Force
        # Try Set-MsolUserPassword (MSOnline module) first
        $msolAvail = Get-Module -ListAvailable -Name MSOnline
        if ($msolAvail) {{
            Import-Module MSOnline -ErrorAction SilentlyContinue
            $msolCred = $null  # App-only; MSOnline needs user cred — skip
        }}
        # Use Microsoft.Graph module (Update-MgUser) — app-only supported
        $mgAvail = Get-Module -ListAvailable -Name Microsoft.Graph.Users
        if ($mgAvail) {{
            Import-Module Microsoft.Graph.Users -ErrorAction SilentlyContinue
            Import-Module Microsoft.Graph.Authentication -ErrorAction SilentlyContinue
            $upn = $mbx.UserPrincipalName
            $pwProfile = @{{
                Password = $pw
                ForceChangePasswordNextSignIn = $false
            }}
            Update-MgUser -UserId $upn -PasswordProfile $pwProfile -AccountEnabled:$true -ErrorAction Stop
            Write-Host (ConvertTo-Json -Compress @{{ stage='password'; idx=$idx; smtp=$smtp; ok=$true }})
        }} else {{
            # Fallback: Set-User doesn't set passwords, flag for Graph retry
            Write-Host (ConvertTo-Json -Compress @{{ stage='password'; idx=$idx; smtp=$smtp; ok=$false; error='Microsoft.Graph.Users module not available' }})
        }}
    }} else {{
        Write-Host (ConvertTo-Json -Compress @{{ stage='password'; idx=$idx; smtp=$smtp; ok=$false; error='no password in CSV' }})
    }}
}} catch {{
    Write-Host (ConvertTo-Json -Compress @{{ stage='error'; idx=$idx; smtp=$smtp; ok=$false; error=$_.Exception.Message }})
}}
"""

    finalize = """
# --- Org-wide SMTP AUTH ---
try {
    Set-TransportConfig -SmtpClientAuthenticationDisabled $false -ErrorAction Stop
    Write-Host (ConvertTo-Json -Compress @{ stage='org-smtpauth'; ok=$true })
} catch {
    Write-Host (ConvertTo-Json -Compress @{ stage='org-smtpauth'; ok=$false; error=$_.Exception.Message })
}
Disconnect-ExchangeOnline -Confirm:$false | Out-Null
"""

    # Per-mailbox state tracker
    per: Dict[int, dict] = {
        i: {
            "idx": i,
            "smtp": _normalize_row(a["mailbox"]).get("PrimarySmtpAddress", ""),
            "upn": None,
            "userId": None,
            "owner": a["owner"],
            "password": _normalize_row(a["mailbox"]).get("Password", ""),
            "create": None,
            "fullaccess": None,
            "sendas": None,
            "sendonbehalf": None,
            "signinUnblocked": None,
            "passwordSet": None,
            "error": None,
        }
        for i, a in enumerate(assignments)
    }

    def handle(line: str) -> None:
        line = line.strip()
        if not line:
            return
        if not line.startswith("{"):
            _log(job_id, "raw", line)
            return
        try:
            evt = json.loads(line)
        except json.JSONDecodeError:
            _log(job_id, "raw", line)
            return
        stage = evt.get("stage")
        idx = evt.get("idx")
        if stage == "create" and isinstance(idx, int):
            per[idx]["create"] = "ok" if evt.get("ok") else "fail"
            per[idx]["upn"] = evt.get("upn")
            _log(job_id, "info", f"[{idx+1}/{len(assignments)}] mailbox created — {evt.get('smtp')}")
        elif stage == "fullaccess" and isinstance(idx, int):
            per[idx]["fullaccess"] = "ok"
            _log(job_id, "info", f"[{idx+1}/{len(assignments)}] full access → {per[idx]['owner']}")
        elif stage == "sendas" and isinstance(idx, int):
            per[idx]["sendas"] = "ok"
            _log(job_id, "info", f"[{idx+1}/{len(assignments)}] send as → {per[idx]['owner']}")
        elif stage == "sendonbehalf" and isinstance(idx, int):
            per[idx]["sendonbehalf"] = "ok"
            _log(job_id, "info", f"[{idx+1}/{len(assignments)}] send-on-behalf → {per[idx]['owner']}")
            job["completed"] = job["completed"] + 1
        elif stage == "password" and isinstance(idx, int):
            if evt.get("ok"):
                per[idx]["passwordSet"] = "ok"
                per[idx]["signinUnblocked"] = "ok"
                _log(job_id, "info", f"[{idx+1}/{len(assignments)}] password set & sign-in enabled — {evt.get('smtp')}")
            else:
                err = evt.get("error", "unknown")
                per[idx]["passwordSet"] = "fail"
                _log(job_id, "warning", f"[{idx+1}/{len(assignments)}] password step: {err}")
        elif stage == "error" and isinstance(idx, int):
            per[idx]["error"] = evt.get("error")
            _log(job_id, "error", f"[{idx+1}/{len(assignments)}] {evt.get('error')}")
        elif stage == "org-smtpauth":
            if evt.get("ok"):
                _log(job_id, "info", "Org-wide SMTP AUTH enabled (Turn-off checkbox unchecked)")
            else:
                _log(job_id, "error", f"Org SMTP AUTH update failed: {evt.get('error')}")

    loop = asyncio.get_event_loop()

    def run_pwsh_script(script: str) -> int:
        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".ps1", delete=False, encoding="utf-8"
        ) as f:
            f.write(script)
            script_path = f.name
        try:
            proc = subprocess.Popen(
                ["pwsh", "-NoProfile", "-NonInteractive", "-File", script_path],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
            )
            for line in iter(proc.stdout.readline, ""):
                handle(line)
            proc.wait()
            return proc.returncode
        finally:
            try:
                os.unlink(script_path)
            except OSError:
                pass

    CHUNK_SIZE = 5
    valid_indices = [i for i, a in enumerate(assignments) if per[i]["smtp"]]
    for i, a in enumerate(assignments):
        if not per[i]["smtp"]:
            _log(job_id, "error", f"[{i+1}] CSV row missing PrimarySmtpAddress — skipped")

    stopped = False
    for chunk_start in range(0, len(valid_indices), CHUNK_SIZE):
        chunk = valid_indices[chunk_start:chunk_start + CHUNK_SIZE]

        # Check for pause/stop before starting this chunk
        while True:
            ctl = job.get("control")
            if ctl == "stop":
                stopped = True
                _log(job_id, "warning", "Job stopped by user")
                break
            if ctl == "pause":
                if job["status"] != "paused":
                    job["status"] = "paused"
                    _log(job_id, "info", "Job paused")
                await asyncio.sleep(1)
                continue
            if job["status"] == "paused":
                job["status"] = "running"
                _log(job_id, "info", "Job resumed")
            job["control"] = None
            break

        if stopped:
            break

        actions = "".join(build_action(i, assignments[i]) for i in chunk)
        chunk_script = _connect_block(cert_path, cert_pw, organization, app_id) + actions + "Disconnect-ExchangeOnline -Confirm:$false | Out-Null\n"

        try:
            rc = await loop.run_in_executor(None, run_pwsh_script, chunk_script)
        except Exception as e:
            _log(job_id, "error", f"PowerShell execution error: {e}")
            for i in chunk:
                if not per[i]["error"]:
                    per[i]["error"] = str(e)
            continue

        if rc != 0:
            _log(job_id, "warning", f"PowerShell exited with code {rc} for this batch")

    if not stopped:
        # --- Org-wide SMTP AUTH (run once at the end) ---
        try:
            rc = await loop.run_in_executor(
                None,
                run_pwsh_script,
                _connect_block(cert_path, cert_pw, organization, app_id) + finalize,
            )
            if rc != 0:
                _log(job_id, "warning", f"PowerShell exited with code {rc} during finalize")
        except Exception as e:
            _log(job_id, "error", f"PowerShell execution error during finalize: {e}")
    # ---------------------------------------------------------------------------
    # Graph fallback: if PowerShell password step failed, retry via Graph API.
    # This handles cases where Microsoft.Graph.Users PS module isn't installed
    # on the Render container but the app has User.ReadWrite.All permission.
    # ---------------------------------------------------------------------------
    needs_graph_retry = [
        i for i, p in per.items()
        if p["create"] == "ok" and p["passwordSet"] != "ok"
    ]

    if needs_graph_retry:
        _log(job_id, "info", f"Retrying password via Graph API for {len(needs_graph_retry)} mailbox(es)…")
        try:
            gt = await get_graph_token(tenant, app_id, sess["client_secret"])
        except HTTPException as e:
            _log(job_id, "error", f"Graph token error: {e.detail}")
            gt = None

        if gt:
            for i in needs_graph_retry:
                st = per[i]
                password = _normalize_row(assignments[i]["mailbox"]).get("Password", "")
                user_ref = st["upn"] or st["smtp"]
                if not password:
                    continue

                last_err: Optional[str] = None
                for attempt in range(4):
                    try:
                        async with httpx.AsyncClient(timeout=60) as client:
                            r = await client.patch(
                                f"{GRAPH_BASE}/users/{user_ref}",
                                headers={
                                    "Authorization": f"Bearer {gt}",
                                    "Content-Type": "application/json",
                                },
                                json={
                                    "accountEnabled": True,
                                    "passwordProfile": {
                                        "password": password,
                                        "forceChangePasswordNextSignIn": False,
                                    },
                                },
                            )
                        if r.status_code >= 400:
                            raise HTTPException(r.status_code, r.text)
                        st["passwordSet"] = "ok"
                        st["signinUnblocked"] = "ok"
                        _log(job_id, "info", f"[{i+1}/{len(assignments)}] Graph: password set & sign-in enabled — {st['smtp']}")
                        last_err = None
                        break
                    except Exception as e:
                        last_err = str(e)
                        await asyncio.sleep(2 + attempt * 2)

                if last_err:
                    st["error"] = (st["error"] + " | " if st["error"] else "") + f"Graph: {last_err}"
                    _log(job_id, "error", f"[{i+1}/{len(assignments)}] Graph password failed: {last_err}")

    any_errors = any(per[i].get("error") for i in per)
    if stopped:
        job["status"] = "stopped"
    else:
        job["status"] = "completed_with_errors" if any_errors else "completed"
    job["finished_at"] = datetime.utcnow().isoformat() + "Z"
    job["results"] = [per[i] for i in range(len(assignments))]
    job["control"] = None
    _log(job_id, "info", f"Job finished — status: {job['status']}")
