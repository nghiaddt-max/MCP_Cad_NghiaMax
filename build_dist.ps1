<#
Dong goi MCP_Cad_NghiaMax thanh file zip de dua len GitHub Releases / mang sang may khac.

    powershell -ExecutionPolicy Bypass -File .\build_dist.ps1

Ket qua: dist\MCP_Cad_NghiaMax-<version>-windows.zip gom file .whl, install.ps1, README.md,
scripts\, templates\, SHA256SUMS.txt; kem dist\...zip.sha256 de nguoi nhan doi chieu.
Can Internet lan dau (pip tai setuptools de build wheel).
#>
$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$py = Join-Path $here ".venv\Scripts\python.exe"
if (-not (Test-Path $py)) { $py = "python" }

$name = "MCP_Cad_NghiaMax"
$version = (& $py -c "import sys; sys.path.insert(0, r'$here'); import autocad_mcp; print(autocad_mcp.__version__)").Trim()
$dist = Join-Path $here "dist"
$stage = Join-Path $dist "$name-$version"
if (Test-Path $stage) { Remove-Item -Recurse -Force $stage }
New-Item -ItemType Directory -Force -Path $stage | Out-Null

& $py -m pip wheel $here --no-deps --disable-pip-version-check -w $stage
if ($LASTEXITCODE) { throw "Build wheel that bai" }
# don thu muc tam setuptools sinh ra trong ma nguon
foreach ($d in @("build", "autocad_mcp.egg-info")) {
    $p = Join-Path $here $d
    if (Test-Path $p) { Remove-Item -Recurse -Force $p }
}
Copy-Item (Join-Path $here "install.ps1"), (Join-Path $here "README.md") $stage
foreach ($d in @("scripts", "templates")) {
    New-Item -ItemType Directory -Force -Path (Join-Path $stage $d) | Out-Null
    Get-ChildItem (Join-Path $here $d) -File | Where-Object { $_.Extension -in ".py", ".dwg" } |
        Copy-Item -Destination (Join-Path $stage $d)
}

# checksum tung file trong goi (duong dan tuong doi, dau / nhu sha256sum)
$lines = Get-ChildItem $stage -File -Recurse | Where-Object { $_.Name -ne "SHA256SUMS.txt" } |
    ForEach-Object {
        $rel = $_.FullName.Substring($stage.Length + 1).Replace("\", "/")
        "{0}  {1}" -f (Get-FileHash $_.FullName -Algorithm SHA256).Hash.ToLower(), $rel
    } | Sort-Object { $_.Substring(66) }
[IO.File]::WriteAllLines((Join-Path $stage "SHA256SUMS.txt"), $lines, (New-Object Text.UTF8Encoding $false))

$zip = Join-Path $dist "$name-$version-windows.zip"
if (Test-Path $zip) { Remove-Item -Force $zip }
Compress-Archive -Path (Join-Path $stage "*") -DestinationPath $zip
$zipHash = (Get-FileHash $zip -Algorithm SHA256).Hash.ToLower()
[IO.File]::WriteAllText("$zip.sha256", "$zipHash  $(Split-Path $zip -Leaf)`n", (New-Object Text.UTF8Encoding $false))
Write-Host "Da dong goi: $zip"
Write-Host "SHA256: $zipHash"
Get-ChildItem $stage -Recurse -File | Format-Table FullName, Length -AutoSize
