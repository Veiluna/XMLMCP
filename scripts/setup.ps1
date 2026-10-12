param([string]$Python = 'python', [switch]$CheckOnly)
$ErrorActionPreference = 'Stop'
$xmlmcpSetup = Join-Path $PSScriptRoot 'setup.py'
if ($CheckOnly) {
    & $Python $xmlmcpSetup --check-only
} else {
    & $Python $xmlmcpSetup
}
if ($LASTEXITCODE -ne 0) { throw 'XMLMCP setup failed; see the error above' }

