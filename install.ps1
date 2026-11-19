<#
Cai dat autocad-mcp tren may Windows.

Chay trong thu muc giai nen (cung cho voi file autocad_mcp-*.whl):
    powershell -ExecutionPolicy Bypass -File .\install.ps1

Tuy chon:
    -InstallDir <thu muc>   Noi cai (mac dinh %LOCALAPPDATA%\autocad-mcp)
    -Python <python.exe>    Chi dinh Python neu khong dung py launcher
    -Wheelhouse <thu muc>   Cai offline tu thu muc chua san cac file .whl phu thuoc
    -Force                  Xoa moi truong cu trong InstallDir roi cai lai
#>
param(
    [string]$InstallDir = (Join-Path $env:LOCALAPPDATA "autocad-mcp"),
    [string]$Python = "",
    [string]$Wheelhouse = "",
    [switch]$Force
)
$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path

$wheel = Get-ChildItem -Path $here -Filter "autocad_mcp-*.whl" | Sort-Object Name | Select-Object -Last 1
if (-not $wheel) { throw "Khong thay file autocad_mcp-*.whl canh install.ps1" }

# --- Python >= 3.10
if (-not $Python) {
    if (Get-Command py -ErrorAction SilentlyContinue) {
        $Python = (& py -3 -c "import sys; print(sys.executable)").Trim()
    } elseif (Get-Command python -ErrorAction SilentlyContinue) {
        $Python = (Get-Command python).Source
    } else {
        throw "Chua cai Python. Tai Python 3.10 tro len tai python.org (tick 'Add python.exe to PATH') roi chay lai."
    }
}
$ver = (& $Python -c "import sys; print('%d.%d' % sys.version_info[:2])").Trim()
if ([version]$ver -lt [version]"3.10") { throw "Can Python 3.10 tro len, dang co $ver ($Python)" }
Write-Host "Python $ver : $Python"

# --- moi truong ao rieng
$venv = Join-Path $InstallDir ".venv"
if (Test-Path $venv) {
    if ($Force) {
        Write-Host "Xoa moi truong cu: $venv"
        Remove-Item -Recurse -Force $venv
    } else {
        Write-Host "Da co $venv - cai de len (dung -Force de cai lai tu dau)"
    }
}
New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
if (-not (Test-Path $venv)) { & $Python -m venv $venv; if ($LASTEXITCODE) { throw "Tao venv that bai" } }
$py = Join-Path $venv "Scripts\python.exe"

# --- cai goi
$pipArgs = @("-m", "pip", "install", "--disable-pip-version-check", $wheel.FullName)
if ($Wheelhouse) { $pipArgs += @("--no-index", "--find-links", $Wheelhouse) }
& $py @pipArgs
if ($LASTEXITCODE) { throw "pip install that bai (kiem tra Internet hoac dung -Wheelhouse)" }

$exe = Join-Path $venv "Scripts\autocad-mcp.exe"
& $exe --version
if ($LASTEXITCODE) { throw "Khong chay duoc $exe" }

# --- file cau hinh MCP mau (UTF-8 khong BOM)
$cfg = [ordered]@{ mcpServers = [ordered]@{ autocad = [ordered]@{ command = $exe; args = @() } } }
$json = $cfg | ConvertTo-Json -Depth 5
$cfgPath = Join-Path $InstallDir "mcp-config.json"
[IO.File]::WriteAllText($cfgPath, $json, (New-Object Text.UTF8Encoding $false))

Write-Host ""
Write-Host "CAI DAT XONG."
Write-Host "  Lenh server : $exe"
Write-Host "  Cau hinh mau: $cfgPath"
Write-Host ""
Write-Host "Kiem tra ket noi (mo AutoCAD va mot ban ve truoc):"
Write-Host "  & `"$exe`" --check"
Write-Host ""
Write-Host "Dang ky voi Claude Code (dung cho moi du an):"
Write-Host "  claude mcp add --scope user autocad -- `"$exe`""
Write-Host "Hoac chep noi dung $cfgPath vao .mcp.json cua du an / claude_desktop_config.json."
