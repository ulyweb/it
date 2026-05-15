<# ============================================
IT Remote Deploy and Network Diagnostics
WPF GUI - Push/Pull files + Network adapter diagnostics
Run As: Administrator
============================================ #>

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="IT Remote Deploy and Network Diagnostics" Height="680" Width="880"
        WindowStartupLocation="CenterScreen" ResizeMode="CanResizeWithGrip"
        Background="#0b1120">
    <Window.Resources>
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="#c9a84c"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="#0f172a"/>
            <Setter Property="Foreground" Value="#e2e8f0"/>
            <Setter Property="BorderBrush" Value="#1e293b"/>
            <Setter Property="FontFamily" Value="Consolas"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Padding" Value="6,4"/>
        </Style>
        <Style TargetType="Button">
            <Setter Property="Background" Value="#0f172a"/>
            <Setter Property="Foreground" Value="#c9a84c"/>
            <Setter Property="BorderBrush" Value="#c9a84c"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Padding" Value="16,8"/>
            <Setter Property="Cursor" Value="Hand"/>
        </Style>
    </Window.Resources>
    <Grid Margin="24">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Grid.Row="0" Text="IT REMOTE DEPLOY AND NETWORK DIAGNOSTICS" FontSize="22" FontWeight="Bold" Margin="0,0,0,2"/>
        <TextBlock Grid.Row="1" Text="Push or pull files to/from remote machines. Run network adapter diagnostics remotely." Foreground="#94a3b8" FontSize="12" Margin="0,0,0,14"/>

        <TextBlock Grid.Row="2" Text="TARGET MACHINE (Hostname or IP)" FontSize="11" FontWeight="Bold" Margin="0,0,0,4"/>
        <TextBox Grid.Row="3" Name="txtTarget" Margin="0,0,0,10"/>

        <StackPanel Grid.Row="4" Orientation="Horizontal" Margin="0,0,0,6">
            <TextBlock Text="TRANSFER MODE:" FontSize="11" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,10,0"/>
            <RadioButton Name="rbPush" Content="PUSH (Local to Remote)" Foreground="#e2e8f0" FontFamily="Segoe UI" FontSize="12" IsChecked="True" VerticalAlignment="Center" Margin="0,0,20,0"/>
            <RadioButton Name="rbPull" Content="PULL (Remote to Local)" Foreground="#e2e8f0" FontFamily="Segoe UI" FontSize="12" VerticalAlignment="Center"/>
        </StackPanel>

        <TextBlock Grid.Row="5" Text="SOURCE FILE OR FOLDER PATH" FontSize="11" FontWeight="Bold" Margin="0,4,0,4"/>
        <Grid Grid.Row="6" Margin="0,0,0,6">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBox Grid.Column="0" Name="txtSource" Margin="0,0,8,0"/>
            <Button Grid.Column="1" Name="btnBrowse" Content="BROWSE" Padding="14,6"/>
        </Grid>

        <TextBlock Grid.Row="6" Text="DESTINATION PATH ON TARGET" FontSize="11" FontWeight="Bold" Margin="0,38,0,4"/>
        <TextBox Grid.Row="7" Name="txtDest" Margin="0,0,0,10" Height="28" VerticalAlignment="Top" Text="C:\Temp"/>

        <ScrollViewer Grid.Row="7" Margin="0,42,0,0" VerticalScrollBarVisibility="Auto" Background="#070d1a" Name="svLog">
            <TextBox Name="txtLog" IsReadOnly="True" TextWrapping="Wrap" Background="#070d1a" Foreground="#38bdf8" BorderThickness="0" VerticalAlignment="Top" FontSize="12"
                     Text="Ready. Enter a target machine and source path, then click TRANSFER NOW or NETWORK INFO."/>
        </ScrollViewer>

        <StackPanel Grid.Row="8" Orientation="Horizontal" HorizontalAlignment="Left" Margin="0,14,0,0">
            <Button Name="btnTransfer" Content="TRANSFER NOW" Margin="0,0,10,0" Padding="20,10"/>
            <Button Name="btnNetInfo" Content="NETWORK INFO" Margin="0,0,10,0" Padding="20,10"/>
            <Button Name="btnExit" Content="EXIT" Foreground="#ef4444" BorderBrush="#ef4444" Padding="20,10"/>
        </StackPanel>
    </Grid>
</Window>
"@

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

$txtTarget   = $window.FindName("txtTarget")
$txtSource   = $window.FindName("txtSource")
$txtDest     = $window.FindName("txtDest")
$txtLog      = $window.FindName("txtLog")
$svLog       = $window.FindName("svLog")
$btnTransfer = $window.FindName("btnTransfer")
$btnNetInfo  = $window.FindName("btnNetInfo")
$btnBrowse   = $window.FindName("btnBrowse")
$btnExit     = $window.FindName("btnExit")
$rbPush      = $window.FindName("rbPush")
$rbPull      = $window.FindName("rbPull")

function Write-Log {
    param([string]$Message)
    $stamp = Get-Date -Format "HH:mm:ss"
    $window.Dispatcher.Invoke([Action]{
        $txtLog.AppendText("`r`n[$stamp] $Message")
        $txtLog.ScrollToEnd()
        $svLog.ScrollToEnd()
    })
}

# --- BROWSE ---
$btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = "Select Source File"
    $dlg.Filter = "All Files (*.*)|*.*"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtSource.Text = $dlg.FileName
    }
})

# --- EXIT ---
$btnExit.Add_Click({ $window.Close() })

# === 3-STAGE CONNECTION TEST (shared by Transfer and Network Info) ===
function Test-RemoteConnection {
    param([string]$Target)

    Write-Log "Target: $Target"

    # Stage 1: Ping
    Write-Log "Stage 1/3: Ping test to $Target..."
    $ping = Test-Connection -ComputerName $Target -Count 1 -Quiet -ErrorAction SilentlyContinue
    if (-not $ping) {
        Write-Log "PING FAILED: $Target is unreachable. Check hostname, network, or firewall."
        return $null
    }
    Write-Log "Ping OK."

    # Stage 2: WinRM
    Write-Log "Stage 2/3: Establishing WinRM session..."
    try {
        $session = New-PSSession -ComputerName $Target -ErrorAction Stop
    } catch {
        Write-Log "WINRM ERROR: $($_.Exception.Message)"
        return $null
    }
    Write-Log "WinRM session established."

    # Stage 3: Admin identity check
    Write-Log "Stage 3/3: Verifying admin identity on remote machine..."
    try {
        $identity = Invoke-Command -Session $session -ScriptBlock {
            $id = [Security.Principal.WindowsIdentity]::GetCurrent()
            $principal = New-Object Security.Principal.WindowsPrincipal($id)
            [pscustomobject]@{
                User = $id.Name
                IsAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
            }
        } -ErrorAction Stop
        Write-Log "Connected as: $($identity.User) | Admin: $($identity.IsAdmin)"
        if (-not $identity.IsAdmin) {
            Write-Log "WARNING: Session is NOT running as Administrator. Some operations may fail."
        }
    } catch {
        Write-Log "IDENTITY CHECK ERROR: $($_.Exception.Message)"
    }

    return $session
}

# === TRANSFER ===
$btnTransfer.Add_Click({
    $target = $txtTarget.Text.Trim()
    $source = $txtSource.Text.Trim()
    $dest   = $txtDest.Text.Trim()
    $isPush = $rbPush.IsChecked

    if ([string]::IsNullOrWhiteSpace($target)) {
        Write-Log "ERROR: Enter a target machine name."
        return
    }
    if ([string]::IsNullOrWhiteSpace($source)) {
        Write-Log "ERROR: Enter a source path."
        return
    }

    $btnTransfer.IsEnabled = $false
    $btnNetInfo.IsEnabled = $false

    $ps = [PowerShell]::Create()
    $ps.AddScript({
        param($target, $source, $dest, $isPush, $syncHash)

        function WL($m) {
            $stamp = Get-Date -Format "HH:mm:ss"
            $syncHash.Window.Dispatcher.Invoke([Action]{
                $syncHash.Log.AppendText("`r`n[$stamp] $m")
                $syncHash.Log.ScrollToEnd()
                $syncHash.Scroll.ScrollToEnd()
            })
        }

        WL "--- FILE TRANSFER INITIATED ---"
        WL "Target: $target"

        # Stage 1: Ping
        WL "Stage 1/2: Ping test to $target..."
        $ping = Test-Connection -ComputerName $target -Count 1 -Quiet -ErrorAction SilentlyContinue
        if (-not $ping) {
            WL "PING FAILED: $target is unreachable."
            $syncHash.Window.Dispatcher.Invoke([Action]{ $syncHash.BtnTransfer.IsEnabled = $true; $syncHash.BtnNetInfo.IsEnabled = $true })
            return
        }
        WL "Ping OK."

        # Stage 2: WinRM
        WL "Stage 2/2: Establishing WinRM session..."
        try {
            $sess = New-PSSession -ComputerName $target -ErrorAction Stop
        } catch {
            WL "WINRM ERROR: $($_.Exception.Message)"
            $syncHash.Window.Dispatcher.Invoke([Action]{ $syncHash.BtnTransfer.IsEnabled = $true; $syncHash.BtnNetInfo.IsEnabled = $true })
            return
        }
        WL "WinRM session established."

        try {
            if ($isPush) {
                WL "Mode: PUSH local -> remote"
                WL "Source: $source"
                WL "Destination: $dest"

                # Ensure destination exists
                Invoke-Command -Session $sess -ArgumentList $dest -ScriptBlock {
                    param($d)
                    if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
                }

                Copy-Item -Path $source -Destination $dest -ToSession $sess -Recurse -Force -ErrorAction Stop
                WL "PUSH complete."

                # SHA-256 verify
                if (Test-Path $source -PathType Leaf) {
                    $localHash = (Get-FileHash -Path $source -Algorithm SHA256).Hash
                    $fileName = [System.IO.Path]::GetFileName($source)
                    $remotePath = "$dest\$fileName"
                    $remoteHash = Invoke-Command -Session $sess -ArgumentList $remotePath -ScriptBlock {
                        param($p)
                        if (Test-Path $p) { (Get-FileHash -Path $p -Algorithm SHA256).Hash } else { "FILE_NOT_FOUND" }
                    }
                    if ($localHash -eq $remoteHash) {
                        WL "SHA-256 VERIFIED: Hashes match."
                    } else {
                        WL "SHA-256 MISMATCH: Local=$localHash Remote=$remoteHash"
                    }
                }
            } else {
                WL "Mode: PULL remote -> local"
                WL "Remote Source: $source"
                WL "Local Destination: $dest"

                if (-not (Test-Path $dest)) { New-Item -ItemType Directory -Path $dest -Force | Out-Null }

                Copy-Item -Path $source -Destination $dest -FromSession $sess -Recurse -Force -ErrorAction Stop
                WL "PULL complete."

                # SHA-256 verify
                $remoteIsFile = Invoke-Command -Session $sess -ArgumentList $source -ScriptBlock {
                    param($p); Test-Path $p -PathType Leaf
                }
                if ($remoteIsFile) {
                    $remoteHash = Invoke-Command -Session $sess -ArgumentList $source -ScriptBlock {
                        param($p); (Get-FileHash -Path $p -Algorithm SHA256).Hash
                    }
                    $fileName = [System.IO.Path]::GetFileName($source)
                    $localPath = Join-Path $dest $fileName
                    if (Test-Path $localPath) {
                        $localHash = (Get-FileHash -Path $localPath -Algorithm SHA256).Hash
                        if ($localHash -eq $remoteHash) {
                            WL "SHA-256 VERIFIED: Hashes match."
                        } else {
                            WL "SHA-256 MISMATCH: Local=$localHash Remote=$remoteHash"
                        }
                    }
                }
            }
        } catch {
            WL "TRANSFER ERROR: $($_.Exception.Message)"
        } finally {
            Remove-PSSession $sess -ErrorAction SilentlyContinue
            WL "Session closed."
            WL "--- FILE TRANSFER COMPLETE ---"
            $syncHash.Window.Dispatcher.Invoke([Action]{ $syncHash.BtnTransfer.IsEnabled = $true; $syncHash.BtnNetInfo.IsEnabled = $true })
        }
    })

    $syncHash = [hashtable]::Synchronized(@{
        Window = $window
        Log = $txtLog
        Scroll = $svLog
        BtnTransfer = $btnTransfer
        BtnNetInfo = $btnNetInfo
    })
    $ps.AddArgument($target).AddArgument($source).AddArgument($dest).AddArgument($isPush).AddArgument($syncHash) | Out-Null
    $ps.BeginInvoke() | Out-Null
})

# === NETWORK INFO ===
$btnNetInfo.Add_Click({
    $target = $txtTarget.Text.Trim()

    if ([string]::IsNullOrWhiteSpace($target)) {
        Write-Log "ERROR: Enter a target machine name."
        return
    }

    $btnTransfer.IsEnabled = $false
    $btnNetInfo.IsEnabled = $false

    $ps = [PowerShell]::Create()
    $ps.AddScript({
        param($target, $syncHash)

        function WL($m) {
            $stamp = Get-Date -Format "HH:mm:ss"
            $syncHash.Window.Dispatcher.Invoke([Action]{
                $syncHash.Log.AppendText("`r`n[$stamp] $m")
                $syncHash.Log.ScrollToEnd()
                $syncHash.Scroll.ScrollToEnd()
            })
        }

        WL "--- NETWORK DIAGNOSTICS INITIATED ---"
        WL "Target: $target"

        # Stage 1: Ping
        WL "Stage 1/2: Ping test to $target..."
        $ping = Test-Connection -ComputerName $target -Count 1 -Quiet -ErrorAction SilentlyContinue
        if (-not $ping) {
            WL "PING FAILED: $target is unreachable. Check hostname, network, or firewall."
            WL "--- NETWORK DIAGNOSTICS COMPLETE ---"
            $syncHash.Window.Dispatcher.Invoke([Action]{ $syncHash.BtnTransfer.IsEnabled = $true; $syncHash.BtnNetInfo.IsEnabled = $true })
            return
        }
        WL "Ping OK."

        # Stage 2: WinRM
        WL "Stage 2/2: Establishing WinRM session..."
        try {
            $sess = New-PSSession -ComputerName $target -ErrorAction Stop
        } catch {
            WL "WINRM ERROR: $($_.Exception.Message)"
            WL "--- NETWORK DIAGNOSTICS COMPLETE ---"
            $syncHash.Window.Dispatcher.Invoke([Action]{ $syncHash.BtnTransfer.IsEnabled = $true; $syncHash.BtnNetInfo.IsEnabled = $true })
            return
        }
        WL "WinRM session established. Querying network configuration..."

        try {
            # Run everything on the remote side and format dates there to avoid deserialization issues
            $netData = Invoke-Command -Session $sess -ScriptBlock {
                $adapters = Get-NetAdapter -ErrorAction SilentlyContinue | ForEach-Object {
                    $a = $_
                    $ipInfo = Get-NetIPAddress -InterfaceIndex $a.ifIndex -ErrorAction SilentlyContinue | Where-Object { $_.AddressFamily -eq 'IPv4' }
                    $dns = Get-DnsClientServerAddress -InterfaceIndex $a.ifIndex -ErrorAction SilentlyContinue | Where-Object { $_.AddressFamily -eq 2 }
                    $gw = Get-NetRoute -InterfaceIndex $a.ifIndex -DestinationPrefix '0.0.0.0/0' -ErrorAction SilentlyContinue | Select-Object -First 1

                    # Format DriverDate on the remote side to avoid ToString deserialization error
                    $driverDateStr = "N/A"
                    if ($null -ne $a.DriverDate) {
                        try { $driverDateStr = "{0:yyyy-MM-dd}" -f $a.DriverDate } catch { $driverDateStr = "N/A" }
                    }

                    $mediaLabel = "Wired"
                    if ($a.MediaType -match "802\.11|WiFi|Wi-Fi|Wireless") { $mediaLabel = "Wireless" }

                    [pscustomobject]@{
                        Name            = $a.Name
                        Description     = $a.InterfaceDescription
                        Status          = [string]$a.Status
                        MediaType       = $mediaLabel
                        LinkSpeed       = $a.LinkSpeed
                        MacAddress      = $a.MacAddress
                        IPv4            = if ($ipInfo) { ($ipInfo.IPAddress -join ", ") } else { "N/A" }
                        SubnetPrefix    = if ($ipInfo) { ($ipInfo.PrefixLength -join ", ") } else { "N/A" }
                        Gateway         = if ($gw) { $gw.NextHop } else { "N/A" }
                        DNS             = if ($dns.ServerAddresses) { ($dns.ServerAddresses -join ", ") } else { "N/A" }
                        DriverVersion   = if ($a.DriverVersion) { $a.DriverVersion } else { "N/A" }
                        DriverDateStr   = $driverDateStr
                        DriverProvider  = if ($a.DriverProvider) { $a.DriverProvider } else { "N/A" }
                        DriverFileName  = if ($a.DriverFileName) { $a.DriverFileName } else { "N/A" }
                    }
                }
                return $adapters
            } -ErrorAction Stop

            if (-not $netData -or @($netData).Count -eq 0) {
                WL "No network adapters found on remote machine."
            } else {
                foreach ($a in @($netData)) {
                    WL "============================================"
                    WL "Adapter: $($a.Name)"
                    WL "  Description:    $($a.Description)"
                    WL "  Status:         $($a.Status)"
                    WL "  Type:           $($a.MediaType)"
                    WL "  Link Speed:     $($a.LinkSpeed)"
                    WL "  MAC Address:    $($a.MacAddress)"
                    WL "  IPv4:           $($a.IPv4)"
                    WL "  Subnet Prefix:  /$($a.SubnetPrefix)"
                    WL "  Gateway:        $($a.Gateway)"
                    WL "  DNS Servers:    $($a.DNS)"
                    WL "  Driver Version: $($a.DriverVersion)"
                    WL "  Driver Date:    $($a.DriverDateStr)"
                    WL "  Driver Provider:$($a.DriverProvider)"
                    WL "  Driver File:    $($a.DriverFileName)"
                }
                WL "============================================"
            }
        } catch {
            WL "QUERY ERROR: $($_.Exception.Message)"
        } finally {
            Remove-PSSession $sess -ErrorAction SilentlyContinue
            WL "Session closed."
            WL "--- NETWORK DIAGNOSTICS COMPLETE ---"
            $syncHash.Window.Dispatcher.Invoke([Action]{ $syncHash.BtnTransfer.IsEnabled = $true; $syncHash.BtnNetInfo.IsEnabled = $true })
        }
    })

    $syncHash = [hashtable]::Synchronized(@{
        Window = $window
        Log = $txtLog
        Scroll = $svLog
        BtnTransfer = $btnTransfer
        BtnNetInfo = $btnNetInfo
    })
    $ps.AddArgument($target).AddArgument($syncHash) | Out-Null
    $ps.BeginInvoke() | Out-Null
})

$window.ShowDialog() | Out-Null
