#Requires -Version 5.1
<#
  PingTray.ps1  v3.1
  Windows tray pinger with Status dashboard, host management GUI,
  Settings dialog (Monitoring/Telegram/SMTP), per-user encrypted config.

  Notes:
  - Comments in English only.
  - Compatible with Windows PowerShell 5.1 and ps2exe.
  - Uses DPAPI via ConvertFrom/To-SecureString for per-user encryption of secrets.
#>

#region ---------- BaseDir & basic config paths ----------
function Get-BaseDir {
    try {
        if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) { return $PSScriptRoot }
        if (-not [string]::IsNullOrWhiteSpace($PSCommandPath)) { return (Split-Path -Parent $PSCommandPath) }
    } catch {}
    return [System.AppDomain]::CurrentDomain.BaseDirectory
}
$Script:BaseDir    = Get-BaseDir
$Script:Version    = '3.1.0.0'
$Script:LogPath    = Join-Path $Script:BaseDir 'PingTray.log'
$Script:HostsFile  = Join-Path $Script:BaseDir 'hosts.txt'
$Script:ConfigFile = Join-Path $Script:BaseDir 'PingTray.config.json'
#endregion

#region ---------- Monitoring defaults (overridden by config if saved) ----------
$Script:IntervalSeconds  = 5
$Script:TimeoutMs        = 1000
$Script:PingBufferSize   = 32
$Script:FailThreshold    = 2
$Script:RecoverThreshold = 2
$Script:AlertCooldownSec = 60
$Script:ShowToasts       = $true
$Script:SoundOnFail      = $false
$Script:BalloonTimeoutMs = 5000
$Script:SmtpTimeoutMs    = 10000
$Script:LogMaxBytes      = 5MB
$Script:DefaultHosts     = @('8.8.8.8','1.1.1.1','karanik.gr')
#endregion

#region ---------- Relaunch in STA when running as .ps1 ----------
$RunningAsPs1 = -not [string]::IsNullOrWhiteSpace($PSCommandPath)
if ($RunningAsPs1 -and $Host.Runspace.ApartmentState -ne 'STA') {
    $psPath = (Get-Process -Id $PID).Path
    $a = @('-NoLogo','-NoProfile','-ExecutionPolicy','Bypass','-STA','-File',"`"$PSCommandPath`"")
    Start-Process -FilePath $psPath -ArgumentList $a -WindowStyle Hidden | Out-Null
    exit
}
#endregion

#region ---------- Assemblies ----------
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class Win32 {
  [DllImport("user32.dll", SetLastError=true)]
  public static extern bool DestroyIcon(IntPtr hIcon);
}
"@
#endregion

#region ---------- Utilities ----------
function Write-Log {
    param([string]$Message)
    $stamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    try {
        if (Test-Path -LiteralPath $Script:LogPath) {
            $fi = New-Object System.IO.FileInfo($Script:LogPath)
            if ($fi.Length -gt $Script:LogMaxBytes) {
                $old = $Script:LogPath + '.old'
                if (Test-Path -LiteralPath $old) { Remove-Item -LiteralPath $old -Force }
                Rename-Item -LiteralPath $Script:LogPath -NewName (Split-Path $old -Leaf) -Force
            }
        }
        Add-Content -Path $Script:LogPath -Value "[$stamp] $Message" -Encoding UTF8
    } catch {}
}

function Protect-Secret {
    param([Parameter(Mandatory=$true)][string]$Plain)
    $sec = ConvertTo-SecureString -String $Plain -AsPlainText -Force
    return ($sec | ConvertFrom-SecureString)
}
function Unprotect-Secret {
    param([Parameter(Mandatory=$true)][string]$Cipher)
    try {
        $sec = ConvertTo-SecureString $Cipher
        $ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)
        try { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr) }
        finally { if ($ptr -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr) } }
    } catch { return '' }
}

function Get-HostsList {
    if (Test-Path -LiteralPath $Script:HostsFile) {
        $lines = Get-Content -LiteralPath $Script:HostsFile -ErrorAction SilentlyContinue
        $h = $lines | ForEach-Object { $_.Trim() } | Where-Object { $_ -and -not $_.StartsWith('#') }
        if ($h.Count -gt 0) { return $h }
    }
    return $Script:DefaultHosts
}
function Save-HostsList {
    $lines = @('# hosts.txt - one host/IP per line','# lines starting with # are comments')
    $lines += $Script:Hosts
    $lines | Set-Content -Path $Script:HostsFile -Encoding UTF8
    Write-Log "Saved hosts.txt: $($Script:Hosts -join ', ')"
}

function Test-Host {
    # Returns hashtable: @{ Success=$true/$false; RoundtripMs=<int or -1> }
    param([string]$HostName)
    try {
        $p = New-Object System.Net.NetworkInformation.Ping
        $buf = New-Object byte[] $Script:PingBufferSize
        $reply = $p.Send($HostName, [int]$Script:TimeoutMs, $buf)
        $p.Dispose()
        $ok = ($reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success)
        $ms = -1; if ($ok) { $ms = $reply.RoundtripTime }
        return @{ Success=$ok; RoundtripMs=$ms }
    } catch { return @{ Success=$false; RoundtripMs=-1 } }
}

function Initialize-HostState {
    $newState = @{}
    foreach ($h in $Script:Hosts) {
        if ($Script:State -and $Script:State.ContainsKey($h)) {
            $newState[$h] = $Script:State[$h]
        } else {
            $newState[$h] = [ordered]@{
                IsDown=$false; FailCount=0; OkCount=0; LastChange=Get-Date
                LastResult=$null; LastAlertTime=[datetime]::MinValue
                LatencyMs=-1; LastDownTime=$null
            }
        }
    }
    $Script:State = $newState
}

function Play-FailSound {
    if (-not $Script:SoundOnFail) { return }
    try {
        # Try playing the wav file directly (works even if sound scheme is "No Sounds")
        $wavPath = "$env:SystemRoot\Media\Windows Exclamation.wav"
        if (Test-Path -LiteralPath $wavPath) {
            $player = New-Object System.Media.SoundPlayer($wavPath)
            $player.Play()
            $player.Dispose()
        } else {
            # Fallback to system sound
            [System.Media.SystemSounds]::Exclamation.Play()
        }
    } catch {
        Write-Log "Play-FailSound error: $($_.Exception.Message)"
    }
}
#endregion

#region ---------- Config load/save ----------
function New-DefaultConfig {
@{
    Monitoring = @{
        IntervalSeconds=5; TimeoutMs=1000; PingBufferSize=32
        FailThreshold=2; RecoverThreshold=2; AlertCooldownSec=60
        ShowToasts=$true; SoundOnFail=$false
    }
    Telegram = @{ Enabled=$false; ChatId=""; TokenEnc="" }
    SMTP     = @{ Enabled=$false; Server=""; Port=587; UseSsl=$true; From=""; User=""; PasswordEnc=""; To="" }
} }

function Load-Config {
    if (Test-Path -LiteralPath $Script:ConfigFile) {
        try {
            $json = Get-Content -LiteralPath $Script:ConfigFile -Raw -Encoding UTF8
            return (ConvertFrom-Json -InputObject $json -ErrorAction Stop)
        } catch { Write-Log "Config load failed: $($_.Exception.Message)" }
    }
    return (New-DefaultConfig)
}
function Save-Config([object]$cfg) {
    try {
        ($cfg | ConvertTo-Json -Depth 6) | Out-File -FilePath $Script:ConfigFile -Encoding UTF8
        Write-Log "Config saved."
        return $true
    } catch { Write-Log "Config save failed: $($_.Exception.Message)"; return $false }
}

$Script:Cfg = Load-Config

# Apply monitoring settings from config (if present)
function Apply-MonitoringConfig {
    $m = $Script:Cfg.Monitoring
    if ($null -eq $m) { return }
    if ($m.IntervalSeconds)  { $Script:IntervalSeconds  = [int]$m.IntervalSeconds }
    if ($m.TimeoutMs)        { $Script:TimeoutMs        = [int]$m.TimeoutMs }
    if ($m.PingBufferSize)   { $Script:PingBufferSize   = [int]$m.PingBufferSize }
    if ($m.FailThreshold)    { $Script:FailThreshold    = [int]$m.FailThreshold }
    if ($m.RecoverThreshold) { $Script:RecoverThreshold = [int]$m.RecoverThreshold }
    if ($m.AlertCooldownSec) { $Script:AlertCooldownSec = [int]$m.AlertCooldownSec }
    if ($null -ne $m.ShowToasts)  { $Script:ShowToasts  = [bool]$m.ShowToasts }
    if ($null -ne $m.SoundOnFail) { $Script:SoundOnFail = [bool]$m.SoundOnFail }
}
Apply-MonitoringConfig
#endregion

#region ---------- Tray icon helpers ----------
function New-StatusIcon {
    param([ValidateSet('Green','Orange','Red')] [string]$ColorName)
    $bmp = New-Object System.Drawing.Bitmap 16,16
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'; $g.Clear([System.Drawing.Color]::Transparent)
    switch ($ColorName) {
        'Green'  { $col = [System.Drawing.Color]::FromArgb(0,166,81) }
        'Orange' { $col = [System.Drawing.Color]::FromArgb(255,140,0) }
        'Red'    { $col = [System.Drawing.Color]::FromArgb(200,0,0) }
    }
    $pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::Black, 1)
    $brush = New-Object System.Drawing.SolidBrush $col
    $g.FillEllipse($brush,1,1,14,14); $g.DrawEllipse($pen,1,1,14,14)
    $hIcon = $bmp.GetHicon()
    $tmpIcon = [System.Drawing.Icon]::FromHandle($hIcon)
    $icon = $tmpIcon.Clone()
    [void][Win32]::DestroyIcon($hIcon)
    $tmpIcon.Dispose(); $pen.Dispose(); $brush.Dispose(); $g.Dispose(); $bmp.Dispose()
    return $icon
}
function Set-TrayStatus {
    param([System.Windows.Forms.NotifyIcon]$NotifyIcon,[ValidateSet('AllUp','PartialDown','AllDown')][string]$Overall)
    switch ($Overall) {
        'AllUp'       { $NotifyIcon.Icon = $Script:IconGreen }
        'PartialDown' { $NotifyIcon.Icon = $Script:IconOrange }
        'AllDown'     { $NotifyIcon.Icon = $Script:IconRed }
    }
    $NotifyIcon.Text = "PingTray v$($Script:Version) - $Overall"
}
$Script:IconGreen  = New-StatusIcon -ColorName Green
$Script:IconOrange = New-StatusIcon -ColorName Orange
$Script:IconRed    = New-StatusIcon -ColorName Red

function Show-Toast {
    param([System.Windows.Forms.NotifyIcon]$NotifyIcon,[string]$Title,[string]$Text)
    if (-not $Script:ShowToasts) { return }
    $NotifyIcon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
    $NotifyIcon.BalloonTipTitle = $Title; $NotifyIcon.BalloonTipText = $Text
    $NotifyIcon.ShowBalloonTip($Script:BalloonTimeoutMs)
}
#endregion

#region ---------- Alert senders ----------
function Send-Telegram {
    param([Parameter(Mandatory=$true)][string]$Text)
    try {
        if (-not $Script:Cfg.Telegram.Enabled) { return }
        $chatId=$Script:Cfg.Telegram.ChatId; $token=$Script:Cfg.Telegram.TokenEnc
        if ([string]::IsNullOrWhiteSpace($chatId) -or [string]::IsNullOrWhiteSpace($token)) { return }
        $tp = Unprotect-Secret $token; if ([string]::IsNullOrWhiteSpace($tp)) { return }
        $body = @{chat_id=$chatId; text=$Text; parse_mode='HTML'; disable_web_page_preview=$true}
        Invoke-WebRequest -Method Post -Uri "https://api.telegram.org/bot$tp/sendMessage" -Body $body -UseBasicParsing -TimeoutSec 10 | Out-Null
        Write-Log "Telegram alert sent."
    } catch { Write-Log "Telegram send failed: $($_.Exception.Message)" }
}
function Send-SMTP {
    param([Parameter(Mandatory=$true)][string]$Subject,[Parameter(Mandatory=$true)][string]$Body)
    try {
        if (-not $Script:Cfg.SMTP.Enabled) { return }
        $svr=$Script:Cfg.SMTP.Server; $port=[int]$Script:Cfg.SMTP.Port; $ssl=[bool]$Script:Cfg.SMTP.UseSsl
        $from=$Script:Cfg.SMTP.From; $to=$Script:Cfg.SMTP.To; $usr=$Script:Cfg.SMTP.User; $pwdE=$Script:Cfg.SMTP.PasswordEnc
        if ([string]::IsNullOrWhiteSpace($svr) -or [string]::IsNullOrWhiteSpace($from) -or [string]::IsNullOrWhiteSpace($to)) { return }
        $msg = New-Object System.Net.Mail.MailMessage; $msg.From=$from
        foreach ($rcpt in $to.Split(',',[System.StringSplitOptions]::RemoveEmptyEntries)) { $msg.To.Add($rcpt.Trim()) }
        $msg.Subject=$Subject; $msg.Body=$Body; $msg.IsBodyHtml=$false
        $cli = New-Object System.Net.Mail.SmtpClient($svr,$port); $cli.EnableSsl=$ssl; $cli.Timeout=$Script:SmtpTimeoutMs
        if (-not [string]::IsNullOrWhiteSpace($usr)) {
            $pwd=""; if (-not [string]::IsNullOrWhiteSpace($pwdE)) { $pwd = Unprotect-Secret $pwdE }
            $cli.Credentials = New-Object System.Net.NetworkCredential($usr,$pwd)
        }
        $cli.Send($msg); $msg.Dispose(); $cli.Dispose()
        Write-Log "SMTP alert sent."
    } catch { Write-Log "SMTP send failed: $($_.Exception.Message)" }
}
#endregion

#region ---------- Settings dialog (3 tabs: Monitoring, Telegram, SMTP) ----------
function Show-SettingsDialog {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "PingTray Settings"; $form.StartPosition = 'CenterScreen'
    $form.Size = New-Object System.Drawing.Size(520, 440)
    $form.FormBorderStyle = 'FixedDialog'; $form.MaximizeBox=$false; $form.MinimizeBox=$false

    $tabs = New-Object System.Windows.Forms.TabControl
    $tabs.Size = New-Object System.Drawing.Size(490, 340); $tabs.Location = New-Object System.Drawing.Point(10,10)
    $tabMon  = New-Object System.Windows.Forms.TabPage; $tabMon.Text = "Monitoring"
    $tabTel  = New-Object System.Windows.Forms.TabPage; $tabTel.Text = "Telegram"
    $tabSmtp = New-Object System.Windows.Forms.TabPage; $tabSmtp.Text = "SMTP"
    $tabs.TabPages.AddRange(@($tabMon,$tabTel,$tabSmtp)); $form.Controls.Add($tabs)

    # ===== Monitoring tab =====
    $tt = New-Object System.Windows.Forms.ToolTip
    $tt.AutoPopDelay = 10000; $tt.InitialDelay = 300; $tt.ReshowDelay = 200

    $y = 15; $lx = 15; $tx = 220; $tw = 100
    $lblInt   = New-Object System.Windows.Forms.Label; $lblInt.Text="Ping Interval (sec):";   $lblInt.Location="$lx,$y";   $lblInt.AutoSize=$true
    $txtInt   = New-Object System.Windows.Forms.TextBox; $txtInt.Location="$tx,$y"; $txtInt.Width=$tw; $txtInt.Text="$($Script:IntervalSeconds)"
    $tt.SetToolTip($lblInt, "How often to ping all hosts (in seconds).`nLower = faster detection, higher CPU/network use.")
    $tt.SetToolTip($txtInt, "How often to ping all hosts (in seconds).")
    $y+=35
    $lblTmo   = New-Object System.Windows.Forms.Label; $lblTmo.Text="Timeout (ms):";          $lblTmo.Location="$lx,$y";   $lblTmo.AutoSize=$true
    $txtTmo   = New-Object System.Windows.Forms.TextBox; $txtTmo.Location="$tx,$y"; $txtTmo.Width=$tw; $txtTmo.Text="$($Script:TimeoutMs)"
    $tt.SetToolTip($lblTmo, "Max time to wait for a ping reply (in milliseconds).`nIf no reply within this time, the host is counted as failed.")
    $tt.SetToolTip($txtTmo, "Max wait time per ping (ms). Default: 1000.")
    $y+=35
    $lblBuf   = New-Object System.Windows.Forms.Label; $lblBuf.Text="Ping Buffer Size (bytes):"; $lblBuf.Location="$lx,$y"; $lblBuf.AutoSize=$true
    $txtBuf   = New-Object System.Windows.Forms.TextBox; $txtBuf.Location="$tx,$y"; $txtBuf.Width=$tw; $txtBuf.Text="$($Script:PingBufferSize)"
    $tt.SetToolTip($lblBuf, "Size of the ICMP payload in bytes.`nDefault: 32. Increase to test with larger packets (e.g. 1472 for MTU testing).")
    $tt.SetToolTip($txtBuf, "ICMP payload size (bytes). Default: 32.")
    $y+=35
    $lblFail  = New-Object System.Windows.Forms.Label; $lblFail.Text="Fail Threshold:";       $lblFail.Location="$lx,$y";  $lblFail.AutoSize=$true
    $txtFail  = New-Object System.Windows.Forms.TextBox; $txtFail.Location="$tx,$y"; $txtFail.Width=$tw; $txtFail.Text="$($Script:FailThreshold)"
    $tt.SetToolTip($lblFail, "Number of consecutive ping failures before a host is marked DOWN.`nHigher = fewer false alarms, slower detection.")
    $tt.SetToolTip($txtFail, "Consecutive failures before DOWN alert. Default: 2.")
    $y+=35
    $lblRec   = New-Object System.Windows.Forms.Label; $lblRec.Text="Recover Threshold:";     $lblRec.Location="$lx,$y";   $lblRec.AutoSize=$true
    $txtRec   = New-Object System.Windows.Forms.TextBox; $txtRec.Location="$tx,$y"; $txtRec.Width=$tw; $txtRec.Text="$($Script:RecoverThreshold)"
    $tt.SetToolTip($lblRec, "Number of consecutive successful pings before a DOWN host is marked UP again.`nPrevents premature recovery alerts from single lucky replies.")
    $tt.SetToolTip($txtRec, "Consecutive successes before RECOVERED alert. Default: 2.")
    $y+=35
    $lblCool  = New-Object System.Windows.Forms.Label; $lblCool.Text="Alert Cooldown (sec):"; $lblCool.Location="$lx,$y";  $lblCool.AutoSize=$true
    $txtCool  = New-Object System.Windows.Forms.TextBox; $txtCool.Location="$tx,$y"; $txtCool.Width=$tw; $txtCool.Text="$($Script:AlertCooldownSec)"
    $tt.SetToolTip($lblCool, "Minimum seconds between Telegram/SMTP alerts for the same host.`nPrevents notification floods when a host bounces up and down rapidly.")
    $tt.SetToolTip($txtCool, "Min seconds between external alerts per host. Default: 60.")
    $y+=35
    $chkToast = New-Object System.Windows.Forms.CheckBox; $chkToast.Text="Show balloon notifications"; $chkToast.Location="$lx,$y"; $chkToast.AutoSize=$true; $chkToast.Checked=$Script:ShowToasts
    $tt.SetToolTip($chkToast, "Show Windows balloon (toast) notifications when a host goes DOWN or recovers.`nUncheck to disable all popup notifications (Telegram/SMTP still work).")
    $y+=30
    $chkSound = New-Object System.Windows.Forms.CheckBox; $chkSound.Text="Play sound on host failure"; $chkSound.Location="$lx,$y"; $chkSound.AutoSize=$true; $chkSound.Checked=$Script:SoundOnFail
    $tt.SetToolTip($chkSound, "Play the Windows Exclamation system sound when a host transitions to DOWN.`nUseful for audible alerting when you're not looking at the screen.")

    $tabMon.Controls.AddRange(@($lblInt,$txtInt,$lblTmo,$txtTmo,$lblBuf,$txtBuf,$lblFail,$txtFail,$lblRec,$txtRec,$lblCool,$txtCool,$chkToast,$chkSound))

    # ===== Telegram tab =====
    $chkTel = New-Object System.Windows.Forms.CheckBox; $chkTel.Text="Enable Telegram alerts"; $chkTel.Location='15,15'; $chkTel.AutoSize=$true
    $lblChat= New-Object System.Windows.Forms.Label; $lblChat.Text="Chat ID:"; $lblChat.Location='15,55'; $lblChat.AutoSize=$true
    $txtChat= New-Object System.Windows.Forms.TextBox; $txtChat.Location='120,50'; $txtChat.Width=320
    $lblTok = New-Object System.Windows.Forms.Label; $lblTok.Text="Bot Token:"; $lblTok.Location='15,90'; $lblTok.AutoSize=$true
    $txtTok = New-Object System.Windows.Forms.TextBox; $txtTok.Location='120,85'; $txtTok.Width=320; $txtTok.UseSystemPasswordChar=$true
    $btnTel = New-Object System.Windows.Forms.Button; $btnTel.Text="Test Telegram"; $btnTel.Location='120,120'; $btnTel.Width=150
    $tabTel.Controls.AddRange(@($chkTel,$lblChat,$txtChat,$lblTok,$txtTok,$btnTel))

    # ===== SMTP tab =====
    $chkS   = New-Object System.Windows.Forms.CheckBox; $chkS.Text="Enable SMTP alerts"; $chkS.Location='15,15'; $chkS.AutoSize=$true
    $lblSrv = New-Object System.Windows.Forms.Label; $lblSrv.Text="Server:"; $lblSrv.Location='15,55'; $lblSrv.AutoSize=$true
    $txtSrv = New-Object System.Windows.Forms.TextBox; $txtSrv.Location='120,50'; $txtSrv.Width=320
    $lblPort= New-Object System.Windows.Forms.Label; $lblPort.Text="Port:"; $lblPort.Location='15,90'; $lblPort.AutoSize=$true
    $txtPort= New-Object System.Windows.Forms.TextBox; $txtPort.Location='120,85'; $txtPort.Width=80
    $chkSSL = New-Object System.Windows.Forms.CheckBox; $chkSSL.Text="Use SSL/TLS"; $chkSSL.Location='220,87'; $chkSSL.AutoSize=$true; $chkSSL.Checked=$true
    $lblFrom= New-Object System.Windows.Forms.Label; $lblFrom.Text="From:"; $lblFrom.Location='15,125'; $lblFrom.AutoSize=$true
    $txtFrom= New-Object System.Windows.Forms.TextBox; $txtFrom.Location='120,120'; $txtFrom.Width=320
    $lblTo  = New-Object System.Windows.Forms.Label; $lblTo.Text="To (comma sep.):"; $lblTo.Location='15,160'; $lblTo.AutoSize=$true
    $txtTo  = New-Object System.Windows.Forms.TextBox; $txtTo.Location='140,155'; $txtTo.Width=300
    $lblUser= New-Object System.Windows.Forms.Label; $lblUser.Text="User:"; $lblUser.Location='15,195'; $lblUser.AutoSize=$true
    $txtUser= New-Object System.Windows.Forms.TextBox; $txtUser.Location='120,190'; $txtUser.Width=320
    $lblPwd = New-Object System.Windows.Forms.Label; $lblPwd.Text="Password:"; $lblPwd.Location='15,230'; $lblPwd.AutoSize=$true
    $txtPwd = New-Object System.Windows.Forms.TextBox; $txtPwd.Location='120,225'; $txtPwd.Width=320; $txtPwd.UseSystemPasswordChar=$true
    $btnSmtp= New-Object System.Windows.Forms.Button; $btnSmtp.Text="Test SMTP"; $btnSmtp.Location='120,260'; $btnSmtp.Width=150
    $tabSmtp.Controls.AddRange(@($chkS,$lblSrv,$txtSrv,$lblPort,$txtPort,$chkSSL,$lblFrom,$txtFrom,$lblTo,$txtTo,$lblUser,$txtUser,$lblPwd,$txtPwd,$btnSmtp))

    # ===== Save / Cancel =====
    $btnSave   = New-Object System.Windows.Forms.Button; $btnSave.Text="Save"; $btnSave.Location='310,360'; $btnSave.Width=80
    $btnCancel = New-Object System.Windows.Forms.Button; $btnCancel.Text="Cancel"; $btnCancel.Location='400,360'; $btnCancel.Width=80
    $form.Controls.AddRange(@($btnSave,$btnCancel))

    # Load config values
    try {
        $chkTel.Checked=[bool]$Script:Cfg.Telegram.Enabled; $txtChat.Text=[string]$Script:Cfg.Telegram.ChatId
        $txtTok.Text=""; if ($Script:Cfg.Telegram.TokenEnc) { $txtTok.Text = Unprotect-Secret $Script:Cfg.Telegram.TokenEnc }
        $chkS.Checked=[bool]$Script:Cfg.SMTP.Enabled; $txtSrv.Text=[string]$Script:Cfg.SMTP.Server
        $txtPort.Text=[string]$Script:Cfg.SMTP.Port; $chkSSL.Checked=[bool]$Script:Cfg.SMTP.UseSsl
        $txtFrom.Text=[string]$Script:Cfg.SMTP.From; $txtTo.Text=[string]$Script:Cfg.SMTP.To
        $txtUser.Text=[string]$Script:Cfg.SMTP.User
        $txtPwd.Text=""; if ($Script:Cfg.SMTP.PasswordEnc) { $txtPwd.Text = Unprotect-Secret $Script:Cfg.SMTP.PasswordEnc }
    } catch {}

    # Test buttons
    $btnTel.Add_Click({
        try {
            $tok=$txtTok.Text.Trim(); $chat=$txtChat.Text.Trim()
            if ([string]::IsNullOrWhiteSpace($tok) -or [string]::IsNullOrWhiteSpace($chat)) {
                [System.Windows.Forms.MessageBox]::Show("Fill both Chat ID and Bot Token.","PingTray") | Out-Null; return }
            Invoke-WebRequest -Uri "https://api.telegram.org/bot$tok/sendMessage" -Method Post -Body @{chat_id=$chat; text="PingTray test OK"} -UseBasicParsing -TimeoutSec 10 | Out-Null
            [System.Windows.Forms.MessageBox]::Show("Telegram test sent.","PingTray") | Out-Null
        } catch { [System.Windows.Forms.MessageBox]::Show("Failed: $($_.Exception.Message)","PingTray") | Out-Null }
    })
    $btnSmtp.Add_Click({
        try {
            $cli = New-Object System.Net.Mail.SmtpClient($txtSrv.Text.Trim(),[int]$txtPort.Text.Trim())
            $cli.EnableSsl=$chkSSL.Checked; $cli.Timeout=$Script:SmtpTimeoutMs
            if (-not [string]::IsNullOrWhiteSpace($txtUser.Text)) { $cli.Credentials = New-Object System.Net.NetworkCredential($txtUser.Text.Trim(),$txtPwd.Text) }
            $msg = New-Object System.Net.Mail.MailMessage; $msg.From=$txtFrom.Text.Trim()
            foreach ($r in $txtTo.Text.Split(',',[System.StringSplitOptions]::RemoveEmptyEntries)) { $msg.To.Add($r.Trim()) }
            $msg.Subject="PingTray SMTP Test"; $msg.Body="Test OK"
            $cli.Send($msg); $msg.Dispose(); $cli.Dispose()
            [System.Windows.Forms.MessageBox]::Show("SMTP test sent.","PingTray") | Out-Null
        } catch { [System.Windows.Forms.MessageBox]::Show("Failed: $($_.Exception.Message)","PingTray") | Out-Null }
    })

    # Save all tabs — rebuild config as hashtable to avoid PSCustomObject property issues
    $btnSave.Add_Click({
        $newCfg = @{
            Monitoring = @{
                IntervalSeconds  = [int]$txtInt.Text
                TimeoutMs        = [int]$txtTmo.Text
                PingBufferSize   = [int]$txtBuf.Text
                FailThreshold    = [int]$txtFail.Text
                RecoverThreshold = [int]$txtRec.Text
                AlertCooldownSec = [int]$txtCool.Text
                ShowToasts       = $chkToast.Checked
                SoundOnFail      = $chkSound.Checked
            }
            Telegram = @{
                Enabled  = $chkTel.Checked
                ChatId   = $txtChat.Text.Trim()
                TokenEnc = ""
            }
            SMTP = @{
                Enabled     = $chkS.Checked
                Server      = $txtSrv.Text.Trim()
                Port        = ([int]($txtPort.Text.Trim()))
                UseSsl      = $chkSSL.Checked
                From        = $txtFrom.Text.Trim()
                To          = $txtTo.Text.Trim()
                User        = $txtUser.Text.Trim()
                PasswordEnc = ""
            }
        }
        if (-not [string]::IsNullOrWhiteSpace($txtTok.Text)) { $newCfg.Telegram.TokenEnc = Protect-Secret $txtTok.Text }
        if (-not [string]::IsNullOrWhiteSpace($txtPwd.Text)) { $newCfg.SMTP.PasswordEnc = Protect-Secret $txtPwd.Text }

        if (Save-Config $newCfg) {
            $Script:Cfg = Load-Config
            Apply-MonitoringConfig
            if ($timer) { $timer.Interval = [int]($Script:IntervalSeconds * 1000) }
            [System.Windows.Forms.MessageBox]::Show("Settings saved. Monitoring parameters applied.","PingTray") | Out-Null
            $form.Close()
        } else { [System.Windows.Forms.MessageBox]::Show("Failed to save (see log).","PingTray") | Out-Null }
    })
    $btnCancel.Add_Click({ $form.Close() })
    $form.ShowDialog() | Out-Null
}
#endregion

#region ---------- Status Window ----------
$Script:StatusForm=$null; $Script:StatusLv=$null; $Script:StatusTxt=$null
$Script:StatusBtn=$null; $Script:StatusBtnRm=$null; $Script:StatusBtnEd=$null
$Script:StatusBtnSet=$null; $Script:StatusBtnCsv=$null; $Script:StatusLbl=$null; $Script:StatusTimer2=$null
$Script:ColorUp=[System.Drawing.Color]::FromArgb(220,255,220)
$Script:ColorDown=[System.Drawing.Color]::FromArgb(255,220,220)
$Script:ColorUnknown=[System.Drawing.Color]::FromArgb(245,245,245)

function Update-StatusListView {
    if ($null -eq $Script:StatusLv) { return }

    # Preserve current selection
    $selectedHosts = @{}
    foreach ($si in $Script:StatusLv.SelectedItems) { $selectedHosts[$si.Text] = $true }

    $Script:StatusLv.BeginUpdate()
    $Script:StatusLv.Items.Clear()
    $upCount=0; $downCount=0
    foreach ($h in $Script:Hosts) {
        $st = $Script:State[$h]; if ($null -eq $st) { continue }
        $item = New-Object System.Windows.Forms.ListViewItem($h)
        # Status
        if ($null -eq $st.LastResult) {
            [void]$item.SubItems.Add('...'); $item.BackColor=$Script:ColorUnknown
        } elseif ($st.IsDown) {
            [void]$item.SubItems.Add('DOWN'); $item.BackColor=$Script:ColorDown; $downCount++
        } else {
            [void]$item.SubItems.Add('UP'); $item.BackColor=$Script:ColorUp; $upCount++
        }
        # Latency
        $lat = ""; if ($st.LatencyMs -ge 0) { $lat = "$($st.LatencyMs) ms" }
        [void]$item.SubItems.Add($lat)
        # Last Check
        $lc = ""; if ($null -ne $st.LastResult) { $lc = (Get-Date).ToString('HH:mm:ss') }
        [void]$item.SubItems.Add($lc)
        # Since
        $since = ""
        if ($st.LastChange) {
            $el = (Get-Date) - $st.LastChange
            if     ($el.TotalHours -ge 24)    { $since = ("{0}d {1}h {2}m" -f [int][math]::Floor($el.TotalDays),$el.Hours,$el.Minutes) }
            elseif ($el.TotalMinutes -ge 60)  { $since = ("{0}h {1}m" -f [int][math]::Floor($el.TotalHours),$el.Minutes) }
            elseif ($el.TotalSeconds -ge 60)  { $since = ("{0}m {1}s" -f [int][math]::Floor($el.TotalMinutes),$el.Seconds) }
            else                              { $since = ("{0}s" -f [int]$el.TotalSeconds) }
        }
        [void]$item.SubItems.Add($since)
        # Consecutive
        if ($st.IsDown) { [void]$item.SubItems.Add("$($st.FailCount) FAIL") }
        else            { [void]$item.SubItems.Add("$($st.OkCount) OK") }
        # Last Down
        $ld = ""; if ($st.LastDownTime) { $ld = ([datetime]$st.LastDownTime).ToString('yyyy-MM-dd HH:mm:ss') }
        [void]$item.SubItems.Add($ld)

        # Restore selection
        if ($selectedHosts.ContainsKey($h)) { $item.Selected = $true }

        [void]$Script:StatusLv.Items.Add($item)
    }
    $Script:StatusLv.EndUpdate()
    $total=$Script:Hosts.Count; $pending=$total-$upCount-$downCount
    if ($Script:StatusLbl) {
        $Script:StatusLbl.Text = "Total: $total  |  Up: $upCount  |  Down: $downCount"
        if ($pending -gt 0) { $Script:StatusLbl.Text += "  |  Pending: $pending" }
    }
}

function Show-StatusWindow {
    if ($Script:StatusForm -and -not $Script:StatusForm.IsDisposed) {
        $Script:StatusForm.WindowState='Normal'; $Script:StatusForm.BringToFront(); $Script:StatusForm.Activate(); return
    }
    $frm = New-Object System.Windows.Forms.Form
    $frm.Text = "PingTray v$($Script:Version) - Host Status"
    $frm.Size = New-Object System.Drawing.Size(800,530); $frm.StartPosition='CenterScreen'
    $frm.MinimumSize = New-Object System.Drawing.Size(700,430)

    # ListView
    $Script:StatusLv = New-Object System.Windows.Forms.ListView
    $Script:StatusLv.View='Details'; $Script:StatusLv.FullRowSelect=$true; $Script:StatusLv.GridLines=$true
    $Script:StatusLv.MultiSelect=$true; $Script:StatusLv.Anchor='Top,Left,Right,Bottom'
    $Script:StatusLv.Location = New-Object System.Drawing.Point(10,10)
    $Script:StatusLv.Size = New-Object System.Drawing.Size(765,340)
    $Script:StatusLv.Font = New-Object System.Drawing.Font('Segoe UI',9)
    [void]$Script:StatusLv.Columns.Add('Host',150)
    [void]$Script:StatusLv.Columns.Add('Status',60)
    [void]$Script:StatusLv.Columns.Add('Latency',70)
    [void]$Script:StatusLv.Columns.Add('Last Check',80)
    [void]$Script:StatusLv.Columns.Add('Since',100)
    [void]$Script:StatusLv.Columns.Add('Consecutive',95)
    [void]$Script:StatusLv.Columns.Add('Last Down',140)
    $frm.Controls.Add($Script:StatusLv)

    # Row 1: Add Host
    $lblAdd = New-Object System.Windows.Forms.Label; $lblAdd.Text="Add Host:"; $lblAdd.Location='10,365'; $lblAdd.AutoSize=$true; $lblAdd.Anchor='Bottom,Left'
    $Script:StatusTxt = New-Object System.Windows.Forms.TextBox; $Script:StatusTxt.Location='80,362'; $Script:StatusTxt.Width=280; $Script:StatusTxt.Anchor='Bottom,Left'
    $Script:StatusTxt.Font = New-Object System.Drawing.Font('Segoe UI',9)
    $Script:StatusBtn = New-Object System.Windows.Forms.Button; $Script:StatusBtn.Text="Add"; $Script:StatusBtn.Location='370,360'; $Script:StatusBtn.Width=60; $Script:StatusBtn.Anchor='Bottom,Left'
    $frm.Controls.AddRange(@($lblAdd,$Script:StatusTxt,$Script:StatusBtn))

    # Row 2: Remove, Edit Hosts, Settings, Export CSV
    $Script:StatusBtnRm  = New-Object System.Windows.Forms.Button; $Script:StatusBtnRm.Text="Remove Selected"; $Script:StatusBtnRm.Location='10,395'; $Script:StatusBtnRm.Width=130; $Script:StatusBtnRm.Anchor='Bottom,Left'
    $Script:StatusBtnEd  = New-Object System.Windows.Forms.Button; $Script:StatusBtnEd.Text="Edit Hosts"; $Script:StatusBtnEd.Location='150,395'; $Script:StatusBtnEd.Width=90; $Script:StatusBtnEd.Anchor='Bottom,Left'
    $Script:StatusBtnSet = New-Object System.Windows.Forms.Button; $Script:StatusBtnSet.Text="Settings"; $Script:StatusBtnSet.Location='250,395'; $Script:StatusBtnSet.Width=80; $Script:StatusBtnSet.Anchor='Bottom,Left'
    $Script:StatusBtnCsv = New-Object System.Windows.Forms.Button; $Script:StatusBtnCsv.Text="Export CSV"; $Script:StatusBtnCsv.Location='340,395'; $Script:StatusBtnCsv.Width=90; $Script:StatusBtnCsv.Anchor='Bottom,Left'
    $frm.Controls.AddRange(@($Script:StatusBtnRm,$Script:StatusBtnEd,$Script:StatusBtnSet,$Script:StatusBtnCsv))

    # Status label
    $Script:StatusLbl = New-Object System.Windows.Forms.Label; $Script:StatusLbl.Text=""; $Script:StatusLbl.Location='10,430'
    $Script:StatusLbl.AutoSize=$true; $Script:StatusLbl.Anchor='Bottom,Left'; $Script:StatusLbl.ForeColor=[System.Drawing.Color]::Gray
    $frm.Controls.Add($Script:StatusLbl)

    Update-StatusListView

    # Auto-refresh timer
    $Script:StatusTimer2 = New-Object System.Windows.Forms.Timer
    $Script:StatusTimer2.Interval = [int]($Script:IntervalSeconds * 1000)
    $Script:StatusTimer2.Add_Tick({ Update-StatusListView })
    $Script:StatusTimer2.Start()

    $frm.Add_FormClosing({
        if ($Script:StatusTimer2) { $Script:StatusTimer2.Stop(); $Script:StatusTimer2.Dispose(); $Script:StatusTimer2=$null }
        $Script:StatusLv=$null; $Script:StatusTxt=$null; $Script:StatusBtn=$null
        $Script:StatusBtnRm=$null; $Script:StatusBtnEd=$null; $Script:StatusBtnSet=$null
        $Script:StatusBtnCsv=$null; $Script:StatusLbl=$null; $Script:StatusForm=$null
    })

    # Add host
    $Script:StatusBtn.Add_Click({
        $newHost = $Script:StatusTxt.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($newHost)) { [System.Windows.Forms.MessageBox]::Show("Enter a hostname or IP.","PingTray") | Out-Null; return }
        $dup=$false; foreach ($ex in $Script:Hosts) { if ($ex -eq $newHost) { $dup=$true; break } }
        if ($dup) { [System.Windows.Forms.MessageBox]::Show("'$newHost' already in list.","PingTray") | Out-Null; return }
        $Script:Hosts += $newHost
        $Script:State[$newHost] = [ordered]@{IsDown=$false;FailCount=0;OkCount=0;LastChange=Get-Date;LastResult=$null;LastAlertTime=[datetime]::MinValue;LatencyMs=-1;LastDownTime=$null}
        Save-HostsList; Write-Log "Host added: $newHost"; $Script:StatusTxt.Text=""; Update-StatusListView
    })
    $Script:StatusTxt.Add_KeyDown({ if ($_.KeyCode -eq 'Return') { $Script:StatusBtn.PerformClick(); $_.SuppressKeyPress=$true } })

    # Remove selected (multi)
    $Script:StatusBtnRm.Add_Click({
        if ($Script:StatusLv.SelectedItems.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Select hosts to remove.","PingTray") | Out-Null; return }
        $rem=@(); foreach ($si in $Script:StatusLv.SelectedItems) { $rem += $si.Text }
        $confirm = [System.Windows.Forms.MessageBox]::Show("Remove?`n`n$($rem -join ', ')","PingTray",'YesNo','Question')
        if ($confirm -ne 'Yes') { return }
        foreach ($r in $rem) { $Script:Hosts=@($Script:Hosts | Where-Object {$_ -ne $r}); if ($Script:State.ContainsKey($r)){$Script:State.Remove($r)}; Write-Log "Host removed: $r" }
        Save-HostsList; Update-StatusListView
    })

    # Edit hosts.txt
    $Script:StatusBtnEd.Add_Click({
        if (-not (Test-Path -LiteralPath $Script:HostsFile)) { Save-HostsList }
        Start-Process notepad.exe $Script:HostsFile -Wait
        $Script:Hosts = Get-HostsList; Initialize-HostState; Write-Log "Hosts reloaded after edit."; Update-StatusListView
    })

    # Settings
    $Script:StatusBtnSet.Add_Click({ Show-SettingsDialog })

    # Export CSV
    $Script:StatusBtnCsv.Add_Click({
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "CSV files (*.csv)|*.csv"; $sfd.FileName = "PingTray_Status_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $sfd.Title = "Export Status to CSV"
        if ($sfd.ShowDialog() -eq 'OK') {
            $rows = @()
            foreach ($h in $Script:Hosts) {
                $st = $Script:State[$h]; if ($null -eq $st) { continue }
                $status = "Pending"; if ($null -ne $st.LastResult) { if ($st.IsDown) { $status="DOWN" } else { $status="UP" } }
                $ld = ""; if ($st.LastDownTime) { $ld = ([datetime]$st.LastDownTime).ToString('yyyy-MM-dd HH:mm:ss') }
                $rows += [PSCustomObject]@{
                    Host=$h; Status=$status; LatencyMs=$st.LatencyMs
                    FailCount=$st.FailCount; OkCount=$st.OkCount
                    LastChange=($st.LastChange).ToString('yyyy-MM-dd HH:mm:ss')
                    LastDown=$ld
                }
            }
            $rows | Export-Csv -Path $sfd.FileName -NoTypeInformation -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("Exported to:`n$($sfd.FileName)","PingTray") | Out-Null
            Write-Log "Status exported to $($sfd.FileName)"
        }
    })

    $Script:StatusForm = $frm
    $frm.Show()
}
#endregion

#region ---------- Initialize state & tray menu ----------
$Script:Hosts = Get-HostsList
$Script:State = @{}
Initialize-HostState
Write-Log "Starting PingTray v$($Script:Version). BaseDir=$($Script:BaseDir). Hosts: $($Script:Hosts -join ', ')"

$tray = New-Object System.Windows.Forms.NotifyIcon
$tray.Visible=$true; $tray.Icon=$Script:IconGreen; $tray.Text="PingTray v$($Script:Version) - Initializing"

$menu = New-Object System.Windows.Forms.ContextMenuStrip
$itemStatus   = $menu.Items.Add('Status');   $menu.Items.Add('-') | Out-Null
$itemSettings = $menu.Items.Add('Settings'); $itemPause = $menu.Items.Add('Pause')
$itemOpenLog  = $menu.Items.Add('Open Log'); $itemReload = $menu.Items.Add('Reload Hosts')
$itemEdit     = $menu.Items.Add('Edit hosts.txt'); $menu.Items.Add('-') | Out-Null
$itemExit     = $menu.Items.Add('Exit')
$tray.ContextMenuStrip = $menu

$Script:IsPaused = $false
$itemStatus.Add_Click({ Show-StatusWindow })
$itemSettings.Add_Click({ Show-SettingsDialog })
$itemPause.Add_Click({
    $Script:IsPaused = -not $Script:IsPaused
    if ($Script:IsPaused) { $itemPause.Text='Resume'; $s='Paused' } else { $itemPause.Text='Pause'; $s='Resumed' }
    Write-Log "Monitoring $s by user"; Show-Toast -NotifyIcon $tray -Title 'PingTray' -Text "Monitoring $s"
})
$itemOpenLog.Add_Click({ if (Test-Path $Script:LogPath) { Start-Process $Script:LogPath } else { Show-Toast -NotifyIcon $tray -Title 'PingTray' -Text 'No log yet.' } })
$itemReload.Add_Click({ $Script:Hosts=Get-HostsList; Initialize-HostState; Write-Log "Reloaded hosts."; Show-Toast -NotifyIcon $tray -Title 'PingTray' -Text ("Reloaded: {0}" -f ($Script:Hosts -join ', ')) })
$itemEdit.Add_Click({
    if (-not (Test-Path $Script:HostsFile)) { @('# hosts.txt','8.8.8.8','1.1.1.1','karanik.gr') | Set-Content -Path $Script:HostsFile -Encoding UTF8 }
    Start-Process notepad.exe $Script:HostsFile | Out-Null
})
$itemExit.Add_Click({
    Write-Log "Exiting PingTray by user"; $timer.Stop()
    if ($Script:StatusTimer2) { try{$Script:StatusTimer2.Stop();$Script:StatusTimer2.Dispose()}catch{} }
    if ($Script:StatusForm -and -not $Script:StatusForm.IsDisposed) { try{$Script:StatusForm.Close()}catch{} }
    $tray.Visible=$false; $tray.Dispose()
    try{$Script:IconGreen.Dispose()}catch{}; try{$Script:IconOrange.Dispose()}catch{}; try{$Script:IconRed.Dispose()}catch{}
    [System.Windows.Forms.Application]::Exit()
})
$tray.Add_DoubleClick({ Show-StatusWindow })
#endregion

#region ---------- Timer & monitoring loop ----------
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = [int]($Script:IntervalSeconds * 1000)
$timer.Add_Tick({
    if ($Script:IsPaused) { return }
    $downCount = 0
    foreach ($h in $Script:Hosts) {
        $ping = Test-Host -HostName $h
        $state = $Script:State[$h]; if ($null -eq $state) { continue }

        $result = $ping.Success
        $state.LatencyMs = $ping.RoundtripMs

        if ($result) { $state.OkCount++; $state.FailCount=0 }
        else         { $state.FailCount++; $state.OkCount=0 }
        $state.LastResult = $result

        # UP -> DOWN
        if (-not $state.IsDown -and -not $result -and $state.FailCount -ge $Script:FailThreshold) {
            $state.IsDown=$true; $state.LastChange=Get-Date; $state.LastDownTime=Get-Date
            $msg="Host DOWN: $h"
            Write-Log $msg; Show-Toast -NotifyIcon $tray -Title 'PingTray Alert' -Text $msg
            Play-FailSound
            $sec = ((Get-Date) - $state.LastAlertTime).TotalSeconds
            if ($sec -ge $Script:AlertCooldownSec) {
                Send-Telegram -Text $msg; Send-SMTP -Subject 'PingTray Alert - DOWN' -Body $msg
                $state.LastAlertTime = Get-Date
            }
        }
        # DOWN -> UP
        elseif ($state.IsDown -and $result -and $state.OkCount -ge $Script:RecoverThreshold) {
            $state.IsDown=$false; $state.LastChange=Get-Date
            $msg="Host RECOVERED: $h"
            Write-Log $msg; Show-Toast -NotifyIcon $tray -Title 'PingTray Alert' -Text $msg
            $sec = ((Get-Date) - $state.LastAlertTime).TotalSeconds
            if ($sec -ge $Script:AlertCooldownSec) {
                Send-Telegram -Text $msg; Send-SMTP -Subject 'PingTray Alert - RECOVERED' -Body $msg
                $state.LastAlertTime = Get-Date
            }
        }
        if ($state.IsDown) { $downCount++ }
    }
    if     ($downCount -eq 0)                   { Set-TrayStatus -NotifyIcon $tray -Overall 'AllUp' }
    elseif ($downCount -ge $Script:Hosts.Count) { Set-TrayStatus -NotifyIcon $tray -Overall 'AllDown' }
    else                                        { Set-TrayStatus -NotifyIcon $tray -Overall 'PartialDown' }
})
$timer.Start()
Set-TrayStatus -NotifyIcon $tray -Overall 'AllUp'
Show-Toast -NotifyIcon $tray -Title 'PingTray' -Text ("Monitoring started. Hosts: {0}" -f ($Script:Hosts -join ', '))
[System.Windows.Forms.Application]::Run()
#endregion
