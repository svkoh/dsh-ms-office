# sideload.ps1 — 手动注册/移除本加载项（WEF 注册表方式；普通用户权限，无需 UAC）
#
# 通常用不到本脚本：桥接服务（server/server.js）启动时会自动注册加载项，并把「按实际监听端口
# 改写过的清单」写到 $DSH_HOME/dsh-office-for-mso/manifest.xml（不碰源码树里的模板）。
# 本脚本用于：自动注册失败、桥接服务暂时不打算跑、或想手工卸载。
#
# 用法：
#   powershell -ExecutionPolicy Bypass -File sideload.ps1
#   powershell -ExecutionPolicy Bypass -File sideload.ps1 -Remove
#   powershell -ExecutionPolicy Bypass -File sideload.ps1 -Manifest C:\path\to\manifest.xml
param(
  [switch]$Remove,
  [string]$Manifest
)

$ErrorActionPreference = 'Stop'

if (-not $Manifest) {
  # 优先用运行期清单（SourceLocation 已按实际端口改写）；没有才退回源码树里的模板（端口 3000）
  $runtimeDir = if ($env:DSH_HOME) { $env:DSH_HOME } else { Join-Path $env:USERPROFILE '.dsh' }
  $runtimeManifest = Join-Path $runtimeDir 'dsh-office-for-mso\manifest.xml'
  $Manifest = if (Test-Path $runtimeManifest) { $runtimeManifest } else { Join-Path $PSScriptRoot 'manifest.xml' }
}

if (-not (Test-Path $Manifest)) {
  Write-Host "[FAIL] 找不到 manifest.xml: $Manifest" -ForegroundColor Red
  exit 1
}

# 从 manifest.xml 提取加载项 GUID（<Id>）
$content = Get-Content -Raw $Manifest
if ($content -match '<Id>\s*([0-9a-fA-F-]{36})\s*</Id>') {
  $guid = $matches[1]
} else {
  Write-Host "[FAIL] manifest.xml 中未找到 <Id> GUID" -ForegroundColor Red
  exit 1
}

$devKey = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
if (-not (Test-Path $devKey)) { New-Item -Path $devKey -Force | Out-Null }

if ($Remove) {
  Remove-ItemProperty -Path $devKey -Name $guid -ErrorAction SilentlyContinue
  Write-Host "[OK] 已移除加载项注册: $guid" -ForegroundColor Green
  Write-Host "     （请关闭并重新打开 Office 文档后生效）" -ForegroundColor Yellow
  exit 0
}

New-ItemProperty -Path $devKey -Name $guid -Value $Manifest -PropertyType String -Force | Out-Null
$check = (Get-ItemProperty -Path $devKey).$guid
if ($check -eq $Manifest) {
  Write-Host "[OK] 加载项已注册: $guid" -ForegroundColor Green
  Write-Host "     manifest: $Manifest" -ForegroundColor Gray
  Write-Host "请关闭并重新打开 Office 文档（或完全退出 Office 后重开），" -ForegroundColor Yellow
  Write-Host "侧边栏将自动出现「DSH Office 执行器」窗格（显示 已连接：等待 DSH 指令）。" -ForegroundColor Yellow
} else {
  Write-Host "[FAIL] 注册校验失败" -ForegroundColor Red
  exit 1
}
