# Base image already has PowerShell 7.4 installed on Ubuntu 22.04
FROM mcr.microsoft.com/powershell:7.4-ubuntu-22.04

# Install Python 3 and pip (Ubuntu 22.04 ships Python 3.10 by default)
RUN apt-get update && \
    DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends \
        python3 python3-pip ca-certificates && \
    rm -rf /var/lib/apt/lists/* && \
    update-ca-certificates

# Install ExchangeOnlineManagement module for all users
RUN pwsh -NoProfile -Command "Set-PSRepository -Name PSGallery -InstallationPolicy Trusted; \
    Install-Module -Name ExchangeOnlineManagement -RequiredVersion 3.7.2 -Force -AllowClobber -Scope AllUsers"

WORKDIR /app

COPY requirements.txt .
RUN python3 -m pip install --no-cache-dir -r requirements.txt

COPY app.py .

# Render injects PORT — default to 8000 locally
ENV PORT=8000
EXPOSE 8000

CMD ["sh", "-c", "python3 -m uvicorn app:app --host 0.0.0.0 --port ${PORT:-8000}"]
