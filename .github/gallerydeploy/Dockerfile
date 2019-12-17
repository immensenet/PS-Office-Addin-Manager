FROM mcr.microsoft.com/powershell:6.1.0-ubuntu-18.04

RUN apt-get update && apt-get install -y git

COPY Publish-Module.ps1 /Publish-Module.ps1
ENTRYPOINT ["pwsh", "/Publish-Module.ps1"]