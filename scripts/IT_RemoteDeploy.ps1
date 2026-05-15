<# ============================================
IT HealthCheck Remote Deploy
Deploy on remote endpoints
Supports Push (local->remote) and Pull (remote->local)
Run As: Administrator on IT Workstation
============================================ #>

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="IT HealthCheck Remote Deploy"
    Width="780" Height="820"
    WindowStartupLocation="CenterScreen"
    ResizeMode="CanResizeWithGrip"
    Background="#0a0e1a">

    <Window.Resources>
        <Style x:Key="GoldLabel" TargetType="TextBlock">
            <Setter Property="Foreground" Value="#c9a84c"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Margin" Value="0,10,0,4"/>
            <Setter Property="FontFamily" Value="Consolas"/>
        </Style>
        <Style x:Key="InputBox" TargetType="TextBox">
            <Setter Property="Background" Value="#10172a"/>
            <Setter Property="Foreground" Value="#e8d9a8"/>
            <Setter Property="BorderBrush" Value="#1e293b"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="FontFamily" Value="Consolas"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="CaretBrush" Value="#c9a84c"/>
        </Style>
        <Style x:Key="ActionBtn" TargetType="Button">
            <Setter Property="Background" Value="#1a1f35"/>
            <Setter Property="Foreground" Value="#c9a84c"/>
            <Setter Property="BorderBrush" Value="#c9a84c"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="14,8"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="FontFamily" Value="Consolas"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Cursor" Value="Hand"/>
        </Style>
    </Window.Resources>

    <Grid Margin="28,20,28,20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <!-- HEADER -->
        <StackPanel Grid.Row="0" Margin="0,0,0,8">
            <TextBlock Text="IT HealthCheck Remote Deploy"
                       FontFamily="Georgia" FontSize="28" FontWeight="Bold"
                       Foreground="#f5f0e8" Margin="0,0,0,2"/>
            <TextBlock Text="Deploy on remote endpoints"
                       FontFamily="Consolas" FontSize="12"
                       Foreground="#8892b0" Margin="0,0,0,0"/>
            <Border Height="1" Background="#1e293b" Margin="0,12,0,0"/>
        </StackPanel>

        <!-- TARGET MACHINE -->
        <StackPanel Grid.Row="1">
            <TextBlock Text="TARGET MACHINE" Style="{StaticResource GoldLabel}"/>
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBox x:Name="txtTarget" Grid.Column="0" Style="{StaticResource InputBox}"
                         ToolTip="Enter hostname or IP of the remote machine"/>
                <Button x:Name="btnTest" Grid.Column="1" Content="TEST CONNECTION"
                        Style="{StaticResource ActionBtn}" Margin="8,0,0,0"/>
            </Grid>
        </StackPanel>

        <!-- CONNECTION STATUS -->
        <TextBlock x:Name="txtConnStatus" Grid.Row="2"
                   Text="" FontFamily="Consolas" FontSize="11"
                   Foreground="#64748b" Margin="0,4,0,0" TextWrapping="Wrap"/>

        <!-- TRANSFER MODE -->
        <StackPanel Grid.Row="3">
            <TextBlock Text="TRANSFER MODE" Style="{StaticResource GoldLabel}"/>
            <ComboBox x:Name="cboMode" Background="#10172a" Foreground="#0a0e1a"
                      FontFamily="Consolas" FontSize="13" Padding="8,6"
                      BorderBrush="#1e293b" BorderThickness="1" SelectedIndex="0">
                <ComboBoxItem Content="Push to Remote  (Local -> Remote)"/>
                <ComboBoxItem Content="Pull from Remote  (Remote -> Local)"/>
            </ComboBox>
        </StackPanel>

        <!-- SOURCE PATH -->
        <StackPanel Grid.Row="4">
            <TextBlock x:Name="lblSource" Text="SOURCE PATH (Local File)" Style="{StaticResource GoldLabel}"/>
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBox x:Name="txtSource" Grid.Column="0" Style="{StaticResource InputBox}"/>
                <Button x:Name="btnBrowseSource" Grid.Column="1" Content="BROWSE"
                        Style="{StaticResource ActionBtn}" Margin="8,0,0,0"/>
            </Grid>
        </StackPanel>

        <!-- TARGET PATH -->
        <StackPanel Grid.Row="5">
            <TextBlock x:Name="lblDest" Text="TARGET PATH (Remote Folder)" Style="{StaticResource GoldLabel}"/>
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBox x:Name="txtDest" Grid.Column="0" Style="{StaticResource InputBox}"
                         Text="C:\IT_Scripts"/>
                <Button x:Name="btnBrowseDest" Grid.Column="1" Content="BROWSE"
                        Style="{StaticResource ActionBtn}" Margin="8,0,0,0"/>
            </Grid>
        </StackPanel>

        <!-- BUTTONS -->
        <StackPanel Grid.Row="6" Orientation="Horizontal" Margin="0,16,0,0">
            <Button x:Name="btnTransfer" Content="TRANSFER NOW" FontSize="14" Padding="22,10"
                    FontWeight="Bold" FontFamily="Consolas" Cursor="Hand"
                    BorderThickness="0" Margin="0,0,12,0">
                <Button.Background>
                    <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
                        <GradientStop Color="#c9a84c" Offset="0"/>
                        <GradientStop Color="#a8843a" Offset="1"/>
                    </LinearGradientBrush>
                </Button.Background>
                <Button.Foreground>
                    <SolidColorBrush Color="#0a0e1a"/>
                </Button.Foreground>
            </Button>
            <Button x:Name="btnExit" Content="EXIT" FontSize="13" Padding="18,10"
                    Background="#1a1f35" Foreground="#ef4444" BorderBrush="#ef4444"
                    BorderThickness="1" FontWeight="Bold" FontFamily="Consolas" Cursor="Hand"/>
        </StackPanel>

        <!-- STATUS -->
        <TextBlock x:Name="txtStatus" Grid.Row="7"
                   Text="Ready." FontFamily="Consolas" FontSize="12"
                   Foreground="#8892b0" Margin="0,12,0,4" TextWrapping="Wrap"/>

        <!-- LOG OUTPUT -->
        <StackPanel Grid.Row="8" Margin="0,4,0,0">
            <TextBlock Text="TRANSFER LOG" Style="{StaticResource GoldLabel}"/>
        </StackPanel>
        <Border Grid.Row="9" Background="#10172a" BorderBrush="#1e293b"
                BorderThickness="1" CornerRadius="4" Margin="0,4,0,0" Padding="4">
            <ScrollViewer x:Name="svLog" VerticalScrollBarVisibility="Auto">
                <TextBox x:Name="txtLog" Background="Transparent" Foreground="#cbd5e1"
                         FontFamily="Consolas" FontSize="11" IsReadOnly="True"
                         TextWrapping="Wrap" BorderThickness="0" AcceptsReturn="True"/>
            </ScrollViewer>
        </Border>
    </Grid>
</Window>
"@

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

$txtTarget      = $window.FindName("txtTarget")
$btnTest        = $window.FindName("btnTest")
$txtConnStatus  = $window.FindName("txtConnStatus")
$cboMode        = $window.FindName("cboMode")
$lblSource      = $window.FindName("lblSource")
$txtSource      = $window.FindName("txtSource")
$btnBrowseSource= $window.FindName("btnBrowseSource")
$lblDest        = $window.FindName("lblDest")
$txtDest        = $window.FindName("txtDest")
$btnBrowseDest  = $window.FindName("btnBrowseDest")
$btnTransfer    = $window.FindName("btnTransfer")
$btnExit        = $window.FindName("btnExit")
$txtStatus      = $window.FindName("txtStatus")
$txtLog         = $window.FindName("txtLog")
$svLog          = $window.FindName("svLog")

# ---- Helper: Append to log ----
function Write-Log {
    param([string]$Message)
    $stamp = Get-Date -Format "HH:mm:ss"
    $window.Dispatcher.Invoke([Action]{
        $txtLog.AppendText("[$stamp] $Message`r`n")
        $txtLog.ScrollToEnd()
    })
}

function Set-Status {
    param([string]$Message, [string]$Color = "#8892b0")
    $window.Dispatcher.Invoke([Action]{
        $txtStatus.Text = $Message
        $txtStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($Color)
    })
}

# ---- Mode Change Handler ----
$cboMode.Add_SelectionChanged({
    if ($cboMode.SelectedIndex -eq 0) {
        # Push mode
        $lblSource.Text  = "SOURCE PATH (Local File)"
        $lblDest.Text    = "TARGET PATH (Remote Folder)"
        $txtSource.Text  = ""
        $txtDest.Text    = "C:\IT_Scripts"
    } else {
        # Pull mode
        $lblSource.Text  = "SOURCE PATH (Remote File or Wildcard)"
        $lblDest.Text    = "TARGET PATH (Local Folder)"
        $txtSource.Text  = "C:\Temp\*_capture_log_*.csv"
        $txtDest.Text    = "C:\Temp"
    }
})

# ---- Browse Source ----
$btnBrowseSource.Add_Click({
    if ($cboMode.SelectedIndex -eq 0) {
        # Push mode: local file picker
        $result = powershell.exe -STA -NoProfile -Command {
            Add-Type -AssemblyName System.Windows.Forms
            $d = New-Object System.Windows.Forms.OpenFileDialog
            $d.Title = "Select file to push to remote machine"
            $d.Filter = "All Files (*.*)|*.*|PowerShell (*.ps1)|*.ps1|CSV (*.csv)|*.csv|HTML (*.html)|*.html"
            if ($d.ShowDialog() -eq 'OK') { $d.FileName } else { "" }
        }
        if ($result) { $txtSource.Text = $result.Trim() }
    } else {
        # Pull mode: user types remote path (no local browse needed)
        [System.Windows.MessageBox]::Show(
            "For Pull mode, type the remote file path or wildcard pattern directly.`n`nExamples:`n  C:\Temp\*_capture_log_*.csv`n  C:\Temp\*.html`n  C:\Temp\HOSTNAME_20260512_143000.html",
            "Remote Path Help",
            "OK", "Information"
        ) | Out-Null
    }
})

# ---- Browse Dest ----
$btnBrowseDest.Add_Click({
    if ($cboMode.SelectedIndex -eq 1) {
        # Pull mode: local folder picker
        $result = powershell.exe -STA -NoProfile -Command {
            Add-Type -AssemblyName System.Windows.Forms
            $d = New-Object System.Windows.Forms.FolderBrowserDialog
            $d.Description = "Select local folder to save pulled files"
            $d.SelectedPath = "C:\Temp"
            if ($d.ShowDialog() -eq 'OK') { $d.SelectedPath } else { "" }
        }
        if ($result) { $txtDest.Text = $result.Trim() }
    } else {
        # Push mode: user types remote folder (no local browse needed)
        [System.Windows.MessageBox]::Show(
            "For Push mode, type the remote destination folder directly.`n`nExamples:`n  C:\IT_Scripts`n  C:\Temp",
            "Remote Folder Help",
            "OK", "Information"
        ) | Out-Null
    }
})

# ---- Test Connection ----
$btnTest.Add_Click({
    $target = $txtTarget.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($target)) {
        $txtConnStatus.Text = "Enter a target machine name or IP."
        $txtConnStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#ef4444")
        return
    }

    $txtConnStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#f59e0b")
    $btnTest.IsEnabled = $false

    # Stage 1: Ping
    $txtConnStatus.Text = "[1/3] Pinging $target ..."
    $window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)

    $pingOk = Test-Connection -ComputerName $target -Count 1 -Quiet -ErrorAction SilentlyContinue
    if (-not $pingOk) {
        $txtConnStatus.Text = "[FAIL] Ping failed. Machine may be offline or blocking ICMP."
        $txtConnStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#ef4444")
        $btnTest.IsEnabled = $true
        return
    }
    $txtConnStatus.Text = "[1/3] Ping OK.  [2/3] Testing WinRM ..."
    $window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)

    # Stage 2: WinRM
    $winrmOk = $false
    try {
        $testSess = New-PSSession -ComputerName $target -ErrorAction Stop
        $winrmOk = $true
        Remove-PSSession $testSess -ErrorAction SilentlyContinue
    } catch {}

    if (-not $winrmOk) {
        $txtConnStatus.Text = "[FAIL] WinRM connection failed. Verify WinRM is enabled on $target."
        $txtConnStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#ef4444")
        $btnTest.IsEnabled = $true
        return
    }
    $txtConnStatus.Text = "[1/3] Ping OK.  [2/3] WinRM OK.  [3/3] Checking Admin ..."
    $window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)

    # Stage 3: Admin check
    $adminOk = $false
    try {
        $testSess = New-PSSession -ComputerName $target -ErrorAction Stop
        $adminOk = Invoke-Command -Session $testSess -ScriptBlock {
            ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        } -ErrorAction Stop
        Remove-PSSession $testSess -ErrorAction SilentlyContinue
    } catch {}

    if ($adminOk) {
        $txtConnStatus.Text = "[1/3] Ping OK.  [2/3] WinRM OK.  [3/3] Admin OK.  --  Connection Verified."
        $txtConnStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#10b981")
    } else {
        $txtConnStatus.Text = "[1/3] Ping OK.  [2/3] WinRM OK.  [3/3] Admin FAIL.  --  Session lacks admin rights."
        $txtConnStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#f59e0b")
    }
    $btnTest.IsEnabled = $true
})

# ---- Transfer ----
$btnTransfer.Add_Click({
    $target   = $txtTarget.Text.Trim()
    $source   = $txtSource.Text.Trim()
    $dest     = $txtDest.Text.Trim()
    $isPush   = ($cboMode.SelectedIndex -eq 0)

    # Validation
    if ([string]::IsNullOrWhiteSpace($target)) {
        Set-Status "ERROR: Enter a target machine." "#ef4444"; return
    }
    if ([string]::IsNullOrWhiteSpace($source)) {
        Set-Status "ERROR: Enter a source path." "#ef4444"; return
    }
    if ([string]::IsNullOrWhiteSpace($dest)) {
        Set-Status "ERROR: Enter a target path." "#ef4444"; return
    }

    if ($isPush -and -not (Test-Path $source)) {
        Set-Status "ERROR: Local source file not found: $source" "#ef4444"; return
    }

    $btnTransfer.IsEnabled = $false
    $txtLog.Text = ""

    $syncHash = [hashtable]::Synchronized(@{
        Window   = $window
        TxtLog   = $txtLog
        TxtStatus= $txtStatus
        BtnTransfer = $btnTransfer
        Target   = $target
        Source   = $source
        Dest     = $dest
        IsPush   = $isPush
    })

    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.ApartmentState = "STA"
    $runspace.Open()
    $runspace.SessionStateProxy.SetVariable("syncHash", $syncHash)

    $ps = [powershell]::Create().AddScript({
        function Log($msg) {
            $stamp = Get-Date -Format "HH:mm:ss"
            $syncHash.Window.Dispatcher.Invoke([Action]{
                $syncHash.TxtLog.AppendText("[$stamp] $msg`r`n")
                $syncHash.TxtLog.ScrollToEnd()
            })
        }
        function Status($msg, $color) {
            $syncHash.Window.Dispatcher.Invoke([Action]{
                $syncHash.TxtStatus.Text = $msg
                $syncHash.TxtStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($color)
            })
        }
        function Done() {
            $syncHash.Window.Dispatcher.Invoke([Action]{
                $syncHash.BtnTransfer.IsEnabled = $true
            })
        }

        $target = $syncHash.Target
        $source = $syncHash.Source
        $dest   = $syncHash.Dest
        $isPush = $syncHash.IsPush

        try {
            # Establish session
            Log "Establishing WinRM session to $target ..."
            Status "Connecting to $target ..." "#f59e0b"
            $sess = New-PSSession -ComputerName $target -ErrorAction Stop
            Log "Session established."

            if ($isPush) {
                # ---- PUSH MODE ----
                Log "Mode: PUSH (Local -> Remote)"
                Log "Source: $source"
                Log "Destination: $dest"

                # Ensure remote folder exists
                Log "Ensuring remote folder exists: $dest"
                Invoke-Command -Session $sess -ArgumentList $dest -ScriptBlock {
                    param($d)
                    if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
                } -ErrorAction Stop
                Log "Remote folder verified."

                # Copy file
                $fileName = [System.IO.Path]::GetFileName($source)
                Log "Copying $fileName to remote ..."
                Status "Copying $fileName ..." "#f59e0b"
                Copy-Item -Path $source -Destination $dest -ToSession $sess -Force -ErrorAction Stop
                Log "File copied successfully."

                # SHA-256 verify
                Log "Computing local SHA-256 ..."
                $localHash = (Get-FileHash -Path $source -Algorithm SHA256).Hash
                Log "Local  SHA-256: $localHash"

                $remotePath = "$dest\$fileName"
                Log "Computing remote SHA-256 ..."
                $remoteHash = Invoke-Command -Session $sess -ArgumentList $remotePath -ScriptBlock {
                    param($p)
                    (Get-FileHash -Path $p -Algorithm SHA256).Hash
                } -ErrorAction Stop
                Log "Remote SHA-256: $remoteHash"

                if ($localHash -eq $remoteHash) {
                    Log "SHA-256 MATCH -- File integrity verified."
                    Status "Push complete. File verified on $target." "#10b981"
                } else {
                    Log "WARNING: SHA-256 MISMATCH. File may be corrupted."
                    Status "Push complete but SHA-256 mismatch!" "#ef4444"
                }

            } else {
                # ---- PULL MODE ----
                Log "Mode: PULL (Remote -> Local)"
                Log "Remote source: $source"
                Log "Local destination: $dest"

                # Ensure local folder exists
                if (-not (Test-Path $dest)) {
                    Log "Creating local folder: $dest"
                    New-Item -ItemType Directory -Path $dest -Force | Out-Null
                }

                # Resolve remote files (supports wildcards)
                Log "Resolving remote file path(s): $source"
                $remoteFiles = Invoke-Command -Session $sess -ArgumentList $source -ScriptBlock {
                    param($pattern)
                    Get-ChildItem -Path $pattern -File -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName
                } -ErrorAction Stop

                if (-not $remoteFiles -or @($remoteFiles).Count -eq 0) {
                    Log "No files matched the pattern: $source"
                    Status "No matching files found on remote machine." "#ef4444"
                    Remove-PSSession $sess -ErrorAction SilentlyContinue
                    Done
                    return
                }

                $fileList = @($remoteFiles)
                Log "Found $($fileList.Count) file(s) to pull."

                $pulled = 0
                foreach ($rf in $fileList) {
                    $rfName = [System.IO.Path]::GetFileName($rf)
                    Log "Pulling: $rfName ..."
                    Status "Pulling $rfName ($($pulled+1)/$($fileList.Count)) ..." "#f59e0b"

                    Copy-Item -Path $rf -Destination $dest -FromSession $sess -Force -ErrorAction Stop
                    $pulled++

                    $localCopy = Join-Path $dest $rfName
                    if (Test-Path $localCopy) {
                        $localHash = (Get-FileHash -Path $localCopy -Algorithm SHA256).Hash
                        $remoteHash = Invoke-Command -Session $sess -ArgumentList $rf -ScriptBlock {
                            param($p)
                            (Get-FileHash -Path $p -Algorithm SHA256).Hash
                        } -ErrorAction Stop

                        if ($localHash -eq $remoteHash) {
                            Log "  SHA-256 verified: $rfName"
                        } else {
                            Log "  WARNING: SHA-256 mismatch for $rfName"
                        }
                    } else {
                        Log "  WARNING: Local file not found after copy: $rfName"
                    }
                }

                Log "Pull complete. $pulled file(s) transferred."
                Status "Pull complete. $pulled file(s) saved to $dest" "#10b981"
            }

            Remove-PSSession $sess -ErrorAction SilentlyContinue
            Log "Session closed."

        } catch {
            Log "ERROR: $($_.Exception.Message)"
            Status "Transfer failed. See log." "#ef4444"
        }

        Done
    })

    $ps.Runspace = $runspace
    $ps.BeginInvoke() | Out-Null
})

# ---- Exit ----
$btnExit.Add_Click({ $window.Close() })

# ---- Show Window ----
$window.ShowDialog() | Out-Null
