"""
Shared Mailbox Provisioner — backend.

Single-file FastAPI app that:
  - Accepts Microsoft 365 app-registration credentials (client ID/secret + .pfx cert)
  - Lists licensed users via Microsoft Graph
  - Bulk-creates shared mailboxes via Exchange Online PowerShell
  - Grants Full Access + Send-on-Behalf to a chosen licensed user
  - Sets a password and unblocks sign-in on each shared mailbox (via Graph)
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
            # Don't let the loop die on transient errors
            await asyncio.sleep(5)


@asynccontextmanager
async def lifespan(app: FastAPI):
    task = asyncio.create_task(_cleanup_loop())
    try:
        yield
    finally:
        task.cancel()


app = FastAPI(title="Shared Mailbox Provisioner", lifespan=lifespan)

# CORS — restrict to your Netlify origin in production via env var
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


# ---------------------------------------------------------------------------
# NEW: resolve Tenant GUID → .onmicrosoft.com domain for Connect-ExchangeOnline
# ---------------------------------------------------------------------------

async def get_tenant_domain(token: str) -> str:
    """
    Exchange Online's -Organization parameter requires the tenant's primary
    domain name (e.g. contoso.onmicrosoft.com), NOT the directory GUID.
    This function looks it up automatically via Graph so the user never has
    to enter it manually.
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
    # Prefer the default .onmicrosoft.com domain
    for d in domains:
        if d["name"].endswith(".onmicrosoft.com") and d.get("isDefault"):
            return d["name"]
    # Fallback: any .onmicrosoft.com domain
    for d in domains:
        if d["name"].endswith(".onmicrosoft.com"):
            return d["name"]
    raise HTTPException(400, "Could not find .onmicrosoft.com domain for this tenant")


async def graph_patch(token: str, path: str, body: dict) -> None:
    async with httpx.AsyncClient(timeout=60) as client:
        r = await client.patch(
            f"{GRAPH_BASE}{path}",
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            },
            json=body,
        )
    if r.status_code >= 400:
        raise HTTPException(r.status_code, f"Graph PATCH {path}: {r.text}")


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
    # NOTE: -Organization must be the tenant domain (e.g. contoso.onmicrosoft.com),
    # NOT the directory GUID. Pass `organization` (resolved via get_tenant_domain),
    # never the raw tenantId GUID.
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
    # Verify pwsh is callable
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
    Validate credentials end-to-end by:
      1. Acquiring a Graph token (verifies client secret + tenant + app ID)
      2. Resolving the tenant GUID → .onmicrosoft.com domain via Graph
      3. Opening + closing an Exchange Online session (verifies cert + Exchange role)
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

    # Resolve the tenant domain — Exchange Online requires the domain name,
    # not the GUID, in the -Organization parameter.
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
        # Trim very long PowerShell tracebacks
        if len(msg) > 1500:
            msg = msg[:1500] + "…"
        raise HTTPException(401, f"Exchange Online connect failed: {msg}")

    token = secrets.token_urlsafe(32)
    SESSIONS[token] = {
        "tenant_id": tenantId,
        "organization": organization,   # resolved domain, used for EXO connections
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

    # Build assignment list: each mailbox row -> owner UPN
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
def get_job_report(token: str, job_id: str):
    get_session(token)
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(404, "Job not found")
    out = io.StringIO()
    w = csv.writer(out)
    w.writerow(
        [
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
    )
    for r in job.get("results", []):
        w.writerow(
            [
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
        )
    return Response(
        content=out.getvalue(),
        media_type="text/csv",
        headers={"Content-Disposition": f"attachment; filename=mailbox-job-{job_id}.csv"},
    )


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
    organization = sess["organization"]   # domain name, not GUID
    app_id = sess["client_id"]
    secret = sess["client_secret"]

    # Build per-mailbox PowerShell actions
    actions = ""
    for i, a in enumerate(assignments):
        m = _normalize_row(a["mailbox"])
        owner = a["owner"]
        smtp = m["PrimarySmtpAddress"]
        if not smtp:
            _log(job_id, "error", f"[{i+1}] CSV row missing PrimarySmtpAddress — skipped")
            continue
        actions += f"""
$idx = {i}
$smtp = '{ps_escape(smtp)}'
$owner = '{ps_escape(owner)}'
try {{
    # ---- Step 1: create mailbox (skip if it already exists) ----
    $existingMbx = Get-Mailbox -Identity $smtp -ErrorAction SilentlyContinue
    if ($existingMbx) {{
        Write-Host (ConvertTo-Json -Compress @{{ stage='create'; idx=$idx; smtp=$smtp; ok=$true; existed=$true; userId=$existingMbx.ExternalDirectoryObjectId; upn=$existingMbx.UserPrincipalName }})
    }} else {{
        $mbx = New-Mailbox -Shared -Name '{ps_escape(m["Alias"])}' -DisplayName '{ps_escape(m["DisplayName"])}' -Alias '{ps_escape(m["Alias"])}' -PrimarySmtpAddress $smtp -ErrorAction Stop
        Write-Host (ConvertTo-Json -Compress @{{ stage='create'; idx=$idx; smtp=$smtp; ok=$true; existed=$false; userId=$mbx.ExternalDirectoryObjectId; upn=$mbx.UserPrincipalName }})
    }}

    # ---- Step 2: grant Full Access (skip if already granted) ----
    $existingFA = Get-MailboxPermission -Identity $smtp -User $owner -ErrorAction SilentlyContinue | Where-Object {{ ($_.AccessRights -contains 'FullAccess') -and (-not $_.IsInherited) }}
    if ($existingFA) {{
        Write-Host (ConvertTo-Json -Compress @{{ stage='fullaccess'; idx=$idx; smtp=$smtp; ok=$true; existed=$true }})
    }} else {{
        Add-MailboxPermission -Identity $smtp -User $owner -AccessRights FullAccess -InheritanceType All -AutoMapping $true -ErrorAction Stop | Out-Null
        Write-Host (ConvertTo-Json -Compress @{{ stage='fullaccess'; idx=$idx; smtp=$smtp; ok=$true; existed=$false }})
    }}

    # ---- Step 3: grant Send As (skip if already granted) ----
    $existingSA = Get-RecipientPermission -Identity $smtp -Trustee $owner -ErrorAction SilentlyContinue | Where-Object {{ $_.AccessRights -contains 'SendAs' }}
    if ($existingSA) {{
        Write-Host (ConvertTo-Json -Compress @{{ stage='sendas'; idx=$idx; smtp=$smtp; ok=$true; existed=$true }})
    }} else {{
        Add-RecipientPermission -Identity $smtp -Trustee $owner -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
        Write-Host (ConvertTo-Json -Compress @{{ stage='sendas'; idx=$idx; smtp=$smtp; ok=$true; existed=$false }})
    }}

    # ---- Step 4: grant Send-on-Behalf (skip if already granted) ----
    $mbxNow = Get-Mailbox -Identity $smtp
    $alreadyOnBehalf = $false
    if ($mbxNow.GrantSendOnBehalfTo) {{
        foreach ($r in $mbxNow.GrantSendOnBehalfTo) {{
            if ("$r" -like "*$owner*") {{ $alreadyOnBehalf = $true; break }}
        }}
    }}
    if ($alreadyOnBehalf) {{
        Write-Host (ConvertTo-Json -Compress @{{ stage='sendonbehalf'; idx=$idx; smtp=$smtp; ok=$true; existed=$true }})
    }} else {{
        Set-Mailbox -Identity $smtp -GrantSendOnBehalfTo @{{ Add=$owner }} -ErrorAction Stop
        Write-Host (ConvertTo-Json -Compress @{{ stage='sendonbehalf'; idx=$idx; smtp=$smtp; ok=$true; existed=$false }})
    }}
}} catch {{
    Write-Host (ConvertTo-Json -Compress @{{ stage='error'; idx=$idx; smtp=$smtp; ok=$false; error=$_.Exception.Message }})
}}
"""

    finalize = """
try {
    Set-TransportConfig -SmtpClientAuthenticationDisabled $false -ErrorAction Stop
    Write-Host (ConvertTo-Json -Compress @{ stage='org-smtpauth'; ok=$true })
} catch {
    Write-Host (ConvertTo-Json -Compress @{ stage='org-smtpauth'; ok=$false; error=$_.Exception.Message })
}
Disconnect-ExchangeOnline -Confirm:$false | Out-Null
"""

    full_script = _connect_block(cert_path, cert_pw, organization, app_id) + actions + finalize

    # Per-mailbox state tracker
    per: Dict[int, dict] = {
        i: {
            "idx": i,
            "smtp": _normalize_row(a["mailbox"]).get("PrimarySmtpAddress", ""),
            "upn": None,
            "userId": None,
            "owner": a["owner"],
            "create": None,
            "existed": False,
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
            per[idx]["userId"] = evt.get("userId")
            per[idx]["upn"] = evt.get("upn")
            per[idx]["existed"] = bool(evt.get("existed"))
            if evt.get("existed"):
                _log(job_id, "info", f"[{idx+1}/{len(assignments)}] mailbox already exists — {evt.get('smtp')}")
            else:
                _log(job_id, "info", f"[{idx+1}/{len(assignments)}] mailbox created — {evt.get('smtp')}")
        elif stage == "fullaccess" and isinstance(idx, int):
            per[idx]["fullaccess"] = "ok"
            if evt.get("existed"):
                _log(job_id, "info", f"[{idx+1}/{len(assignments)}] full access already set → {per[idx]['owner']}")
            else:
                _log(job_id, "info", f"[{idx+1}/{len(assignments)}] full access → {per[idx]['owner']}")
        elif stage == "sendas" and isinstance(idx, int):
            per[idx]["sendas"] = "ok"
            if evt.get("existed"):
                _log(job_id, "info", f"[{idx+1}/{len(assignments)}] send as already set → {per[idx]['owner']}")
            else:
                _log(job_id, "info", f"[{idx+1}/{len(assignments)}] send as → {per[idx]['owner']}")
        elif stage == "sendonbehalf" and isinstance(idx, int):
            per[idx]["sendonbehalf"] = "ok"
            if evt.get("existed"):
                _log(job_id, "info", f"[{idx+1}/{len(assignments)}] send-on-behalf already set → {per[idx]['owner']}")
            else:
                _log(job_id, "info", f"[{idx+1}/{len(assignments)}] send-on-behalf → {per[idx]['owner']}")
            job["completed"] = job["completed"] + 1
        elif stage == "error" and isinstance(idx, int):
            per[idx]["error"] = evt.get("error")
            _log(job_id, "error", f"[{idx+1}/{len(assignments)}] {evt.get('error')}")
        elif stage == "org-smtpauth":
            if evt.get("ok"):
                _log(job_id, "info", "Org-wide SMTP AUTH enabled (Turn-off checkbox unchecked)")
            else:
                _log(job_id, "error", f"Org SMTP AUTH update failed: {evt.get('error')}")

    loop = asyncio.get_event_loop()

    def run_pwsh_stream() -> int:
        # Write the script to a temp .ps1 file and run it with -File.
        # This avoids the OS "argument list too long" (Errno 7) limit when
        # the script gets large (e.g. 99+ mailboxes × idempotency checks).
        fd, script_path = tempfile.mkstemp(suffix=".ps1", prefix="mp_job_")
        try:
            with os.fdopen(fd, "w") as f:
                f.write(full_script)
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

    try:
        rc = await loop.run_in_executor(None, run_pwsh_stream)
    except Exception as e:
        _log(job_id, "error", f"PowerShell execution error: {e}")
        job["status"] = "failed"
        job["finished_at"] = datetime.utcnow().isoformat() + "Z"
        return

    if rc != 0:
        _log(job_id, "warning", f"PowerShell exited with code {rc}")

    # Graph phase: set password + unblock sign-in
    _log(job_id, "info", "Setting passwords and unblocking sign-in via Graph…")
    try:
        gt = await get_graph_token(tenant, app_id, secret)
    except HTTPException as e:
        _log(job_id, "error", f"Graph token error: {e.detail}")
        job["status"] = "failed"
        job["finished_at"] = datetime.utcnow().isoformat() + "Z"
        return

    for i, a in enumerate(assignments):
        st = per[i]
        if st["create"] != "ok":
            continue
        password = _normalize_row(a["mailbox"]).get("Password", "")
        user_ref = st["userId"] or st["upn"] or st["smtp"]
        if not password:
            _log(job_id, "warning", f"[{i+1}/{len(assignments)}] no password in CSV — skipping")
            continue

        # If the mailbox already existed AND the underlying user account is already
        # enabled, assume the password was already set on a previous run and skip.
        if st.get("existed"):
            try:
                async with httpx.AsyncClient(timeout=30) as client:
                    rr = await client.get(
                        f"{GRAPH_BASE}/users/{user_ref}?$select=accountEnabled",
                        headers={"Authorization": f"Bearer {gt}"},
                    )
                if rr.status_code == 200 and rr.json().get("accountEnabled"):
                    st["signinUnblocked"] = "ok"
                    st["passwordSet"] = "skipped"
                    _log(
                        job_id,
                        "info",
                        f"[{i+1}/{len(assignments)}] sign-in already enabled — {st['smtp']}",
                    )
                    continue
            except Exception:
                # If the check fails, fall through and try the PATCH anyway
                pass

        # Brief retry — newly-created user can take a moment to be visible in Graph
        last_err: Optional[str] = None
        for attempt in range(4):
            try:
                await graph_patch(
                    gt,
                    f"/users/{user_ref}",
                    {
                        "accountEnabled": True,
                        "passwordProfile": {
                            "password": password,
                            "forceChangePasswordNextSignIn": False,
                        },
                    },
                )
                st["signinUnblocked"] = "ok"
                st["passwordSet"] = "ok"
                _log(
                    job_id,
                    "info",
                    f"[{i+1}/{len(assignments)}] sign-in unblocked & password set — {st['smtp']}",
                )
                last_err = None
                break
            except HTTPException as e:
                last_err = str(e.detail)
                await asyncio.sleep(2 + attempt * 2)
        if last_err:
            st["error"] = (st["error"] + " | " if st["error"] else "") + f"Graph: {last_err}"
            _log(job_id, "error", f"[{i+1}/{len(assignments)}] Graph update failed: {last_err}")

    any_errors = any(per[i].get("error") for i in per)
    job["status"] = "completed_with_errors" if any_errors else "completed"
    job["finished_at"] = datetime.utcnow().isoformat() + "Z"
    job["results"] = [per[i] for i in range(len(assignments))]
    _log(job_id, "info", f"Job finished — status: {job['status']}")
