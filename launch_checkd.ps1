param(
  [Parameter(Mandatory = $true)][string]$Token,
  [string]$Config = "",
  [switch]$Observe
)
# Detached launcher for the self-check + auto-retry daemon (headless: no window, no focus).
$root = if ($PSScriptRoot) { $PSScriptRoot } else { "." }   # 脚本与配置所在目录，可改成你的绝对路径
$py = "python"                                              # 也可写死成你的 python.exe 绝对路径
if (-not $Config) { $Config = "$root/tk_publish.config.json" }
$env:PYTHONIOENCODING = "utf-8"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$log = "$root/check_$stamp.log"
$mode = @()
if ($Observe) { $mode = @("--no-retry") }
$p = Start-Process -FilePath $py -WorkingDirectory $root -WindowStyle Hidden -PassThru `
  -ArgumentList (@("-X", "utf8", "$root/tk_publish.py", "checkd", "--token", $Token, "--config", $Config) + $mode) `
  -RedirectStandardOutput $log -RedirectStandardError "$log.err"
Write-Output ("DETACHED PID=" + $p.Id + " LOG=" + $log)
