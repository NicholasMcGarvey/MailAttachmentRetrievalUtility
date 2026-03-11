#Requires -Version 5.1
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms   # for FolderBrowserDialog

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------

$ScriptDir   = Split-Path -Parent $MyInvocation.MyCommand.Path
$ScriptPath  = Join-Path $ScriptDir "MARU.ps1"
$UpdaterPath = Join-Path $ScriptDir "MARU_Update.ps1"
$ConfigPath  = Join-Path $ScriptDir "MARU_Configs.json"
$LastRunPath = Join-Path $ScriptDir "MARU_LastRun.json"

$VersionFileName = "MARU_Version.json"

# ---------------------------------------------------------------------------
# Update source - set this to the UNC version subfolder on your file share.
# e.g. "\\server\share\MARU\version"
# Set to $null or "" to disable update checks.
# ---------------------------------------------------------------------------

$UpdateSourcePath = "S:\Corporate Trust\Corporate Trust\Utilities\Mail Attachment Retrieval Utility\Version"

# ---------------------------------------------------------------------------
# Version helpers
# ---------------------------------------------------------------------------

$ManagedFiles = @("MARU.ps1", "MARU_UI.ps1", "MARU_Update.ps1")

function Get-FileChecksum([string]$Path) {
    return (Get-FileHash -Path $Path -Algorithm SHA256).Hash
}

function Get-VersionManifest([string]$SourceDir) {
    $manifestPath = Join-Path $SourceDir $VersionFileName
    if (-not (Test-Path $manifestPath)) { return $null }
    try { return Get-Content $manifestPath -Raw | ConvertFrom-Json } catch { return $null }
}

function Get-OutdatedFiles([string]$SourceDir) {
    $manifest = Get-VersionManifest $SourceDir
    if ($null -eq $manifest) { return $null }
    $outdated = @()
    foreach ($entry in $manifest.Files) {
        $localPath = Join-Path $ScriptDir $entry.Name
        if (-not (Test-Path $localPath)) { $outdated += $entry.Name; continue }
        if ((Get-FileChecksum $localPath) -ne $entry.Checksum) { $outdated += $entry.Name }
    }
    return [PSCustomObject]@{ Version = $manifest.Version; ReleaseNotes = $manifest.ReleaseNotes; Files = $outdated }
}

function Get-LocalVersion {
    $localManifest = Join-Path $ScriptDir $VersionFileName
    if (Test-Path $localManifest) {
        try { return (Get-Content $localManifest -Raw | ConvertFrom-Json).Version } catch {}
    }
    return "unknown"
}

function Invoke-UpdateCheck([bool]$ManualTrigger = $false) {
    if ([string]::IsNullOrWhiteSpace($UpdateSourcePath)) { return }

    try { $result = Get-OutdatedFiles $UpdateSourcePath }
    catch {
        if ($ManualTrigger) {
            [System.Windows.MessageBox]::Show(
                "Could not reach the update source:`n$UpdateSourcePath`n`nError: $_",
                "Update Check Failed",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning) | Out-Null
        }
        return
    }

    if ($null -eq $result) {
        if ($ManualTrigger) {
            [System.Windows.MessageBox]::Show(
                "Version manifest not found at:`n$UpdateSourcePath",
                "Update Check",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning) | Out-Null
        }
        return
    }

    if ($result.Files.Count -eq 0) {
        if ($ManualTrigger) {
            [System.Windows.MessageBox]::Show(
                "You are up to date (version $($result.Version)).",
                "Update Check",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information) | Out-Null
        }
        return
    }

    $localVer = Get-LocalVersion
    $notes    = if ($result.ReleaseNotes) { "`n`nRelease notes: $($result.ReleaseNotes)" } else { "" }
    $msg      = "An update is available.`n`nInstalled : $localVer`nAvailable : $($result.Version)$notes`n`nFiles to update: $($result.Files -join ", ")`n`nInstall now and relaunch?"

    $choice = [System.Windows.MessageBox]::Show(
        $msg, "Update Available",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question)
    if ($choice -ne [System.Windows.MessageBoxResult]::Yes) { return }

    $filesToUpdateArg = $result.Files -join ","
    Start-Process -FilePath "powershell.exe" `
                  -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$UpdaterPath`" -LocalDir `"$ScriptDir`" -VersionSourceDir `"$UpdateSourcePath`" -FilesToUpdate `"$filesToUpdateArg`"" `
                  -WindowStyle Hidden
    $window.Close()
}

# ---------------------------------------------------------------------------
# Outlook default mailbox
# ---------------------------------------------------------------------------

function Get-DefaultMailBoxName {
    try {
        $ol   = New-Object -ComObject Outlook.Application
        $ns   = $ol.GetNamespace("MAPI")
        $addr = $ns.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress
        if ($addr) { return $addr }
    } catch {}
    try {
        $ol   = New-Object -ComObject Outlook.Application
        $ns   = $ol.GetNamespace("MAPI")
        $addr = $ns.CurrentUser.Address
        if ($addr) { return $addr }
    } catch {}
    return $env:USERNAME
}

# ---------------------------------------------------------------------------
# Folder browser helper (WinForms)
# ---------------------------------------------------------------------------

function Show-FolderPicker([string]$InitialPath = "") {
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Select folder"
    $dlg.ShowNewFolderButton = $true
    if ($InitialPath -and (Test-Path $InitialPath)) { $dlg.SelectedPath = $InitialPath }
    $result = $dlg.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) { return $dlg.SelectedPath }
    return $null
}

# ---------------------------------------------------------------------------
# Config helpers
# ---------------------------------------------------------------------------

function Load-Configs {
    $list = [System.Collections.Generic.List[PSCustomObject]]::new()
    if (Test-Path $ConfigPath) {
        try {
            $items = Get-Content $ConfigPath -Raw | ConvertFrom-Json
            if ($null -ne $items) { foreach ($item in @($items)) { $list.Add($item) } }
        } catch {}
    }
    return $list
}

function Save-Configs($configs) {
    $configs | ConvertTo-Json -Depth 5 | Out-File $ConfigPath -Encoding utf8
}

function Load-LastRun {
    if (Test-Path $LastRunPath) {
        try { return Get-Content $LastRunPath -Raw | ConvertFrom-Json } catch {}
    }
    return $null
}

function Save-LastRun($profile) {
    $profile | ConvertTo-Json -Depth 5 | Out-File $LastRunPath -Encoding utf8
}

function New-EmptyProfile {
    return [PSCustomObject]@{
        ProfileName            = "New Profile"
        MailBoxName            = ""
        MailBoxFolderName      = "Inbox"
        SearchSubFolders       = $false
        FilterSubject          = ""
        FilterSender           = ""
        FilterTo               = ""
        FilterCC               = ""
        FilterBCC              = ""
        SaveToFolders          = ""
        LogPath                = ""
        DaysBack               = ""
        FromDate               = ""
        ToDate                 = ""
        SkipAlreadyDownloaded  = $true
        NoLog                  = $false
        FileCollisionAction    = "Suffix"
        CreateFolderPreference = "Last"
    }
}

function Parse-FilterValues($raw) {
    return ($raw -split "[\r\n,]+") | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
}

function Parse-Folders($raw) { return Parse-FilterValues $raw }

function Test-DateString($value) {
    if ([string]::IsNullOrWhiteSpace($value)) { return $true }
    $parsed = [System.DateTime]::MinValue
    return [System.DateTime]::TryParse(
        $value,
        [System.Globalization.CultureInfo]::CurrentCulture,
        [System.Globalization.DateTimeStyles]::None,
        [ref]$parsed)
}

function Set-DateFieldState($textBox, $isValid) {
    if ($isValid) {
        $textBox.BorderBrush     = [System.Windows.Media.Brushes]::Silver
        $textBox.BorderThickness = [System.Windows.Thickness]::new(1)
        $textBox.ToolTip         = "Accepted formats: yyyy-MM-dd, MM/dd/yyyy, dd MMM yyyy, and most common date formats."
    } else {
        $textBox.BorderBrush     = [System.Windows.Media.SolidColorBrush]::new(
                                       [System.Windows.Media.Color]::FromRgb(192, 57, 43))
        $textBox.BorderThickness = [System.Windows.Thickness]::new(2)
        $textBox.ToolTip         = "Invalid date. Accepted formats: yyyy-MM-dd, MM/dd/yyyy, dd MMM yyyy, etc."
    }
}

# ---------------------------------------------------------------------------
# XAML
# ---------------------------------------------------------------------------

[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Mail Attachment Retrieval Utility" Height="900" Width="1160" MinHeight="640" MinWidth="900"
    WindowStartupLocation="CenterScreen" FontFamily="Segoe UI" FontSize="13"
    Background="#F0F2F5">

    <Window.Resources>
        <Style TargetType="TextBox">
            <Setter Property="Padding" Value="6,4"/>
            <Setter Property="BorderBrush" Value="#CCCCCC"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Background" Value="White"/>
        </Style>
        <Style TargetType="ComboBox">
            <Setter Property="Padding" Value="6,4"/>
            <Setter Property="BorderBrush" Value="#CCCCCC"/>
        </Style>
        <Style x:Key="SidebarButton" TargetType="Button">
            <Setter Property="Padding" Value="6,4"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontSize" Value="12"/>
        </Style>
        <Style x:Key="BrowseButton" TargetType="Button">
            <Setter Property="Padding" Value="6,4"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="#CCCCCC"/>
            <Setter Property="Background" Value="#EEEEEE"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="VerticalAlignment" Value="Bottom"/>
        </Style>
        <Style x:Key="Label" TargetType="TextBlock">
            <Setter Property="Margin" Value="0,6,0,2"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Foreground" Value="#333333"/>
        </Style>
        <Style x:Key="RequiredLabel" TargetType="TextBlock">
            <Setter Property="Margin" Value="0,6,0,2"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Foreground" Value="#333333"/>
        </Style>
        <Style x:Key="SectionHeader" TargetType="TextBlock">
            <Setter Property="FontSize" Value="11"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="Foreground" Value="#888888"/>
            <Setter Property="Margin" Value="0,10,0,0"/>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="175"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="200"/>
            <RowDefinition Height="40"/>
        </Grid.RowDefinitions>

        <!-- Sidebar -->
        <Border Grid.Column="0" Grid.Row="0" Grid.RowSpan="3"
                Background="#2C3E50" Padding="10">
            <DockPanel>
                <TextBlock DockPanel.Dock="Top" Text="PROFILES"
                           Foreground="#95A5A6" FontSize="11" FontWeight="Bold"
                           Margin="0,8,0,10"/>
                <StackPanel DockPanel.Dock="Bottom" Margin="0,8,0,0">
                    <Button x:Name="btnNew"    Content="+ New"  Margin="0,2"
                            Background="#27AE60" Foreground="White"
                            Style="{StaticResource SidebarButton}"/>
                    <Button x:Name="btnSave"   Content="Save"   Margin="0,2"
                            Background="#2980B9" Foreground="White"
                            Style="{StaticResource SidebarButton}"/>
                    <Button x:Name="btnDelete" Content="Delete" Margin="0,2"
                            Background="#E74C3C" Foreground="White"
                            Style="{StaticResource SidebarButton}"/>
                    <Separator Margin="0,6,0,4" Background="#3D5166"/>
                    <Button x:Name="btnCheckUpdate" Content="Check for Updates" Margin="0,2"
                            Background="#566573" Foreground="#BDC3C7"
                            Style="{StaticResource SidebarButton}"
                            ToolTip="Check the file share for a newer version of MARU."/>
                </StackPanel>
                <ListBox x:Name="lstProfiles" Background="Transparent"
                         BorderThickness="0" Foreground="White"
                         HorizontalContentAlignment="Stretch">
                    <ListBox.ItemContainerStyle>
                        <Style TargetType="ListBoxItem">
                            <Setter Property="Padding" Value="8,6"/>
                            <Setter Property="Cursor" Value="Hand"/>
                            <Style.Triggers>
                                <Trigger Property="IsSelected" Value="True">
                                    <Setter Property="Background" Value="#3498DB"/>
                                    <Setter Property="Foreground" Value="White"/>
                                </Trigger>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Setter Property="Background" Value="#34495E"/>
                                </Trigger>
                            </Style.Triggers>
                        </Style>
                    </ListBox.ItemContainerStyle>
                </ListBox>
            </DockPanel>
        </Border>

        <!-- Main form -->
        <ScrollViewer Grid.Column="1" Grid.Row="0"
                      VerticalScrollBarVisibility="Auto" Padding="20,14,20,8">
            <StackPanel>

                <!-- Profile name -->
               
                <!-- PROFILE NAME: label + right-aligned checkbox on the same row,
                     with the textbox on the row below -->
                <Grid Margin="0,0,0,8">
                      <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>  <!-- Row 0: label + checkbox -->
                            <RowDefinition Height="Auto"/>  <!-- Row 1: textbox -->
                      </Grid.RowDefinitions>
                      <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>   <!-- Col 0: label (left) -->
                            <ColumnDefinition Width="Auto"/><!-- Col 1: checkbox (right) -->
                      </Grid.ColumnDefinitions>

                      <!-- Left: Profile Name label (same row as checkbox) -->
                      <TextBlock
                          Grid.Row="0" Grid.Column="0"
                          Style="{StaticResource Label}"
                          Text="Profile Name"
                          VerticalAlignment="Center" />

                      <!-- Right: your checkbox on the same row -->
                      <CheckBox
                          Grid.Row="0" Grid.Column="1"
                          x:Name="chkVerbose"
                          Content="Verbose"
                          VerticalAlignment="Center"
                          HorizontalAlignment="Right"
                          Margin="12,0,0,0"/>

                      <!-- Row below: the Profile Name textbox (full width) -->
                      <TextBox
                          Grid.Row="1" Grid.Column="0" Grid.ColumnSpan="2"
                          x:Name="txtProfileName"/>
                </Grid>


                <!-- ── MAILBOX ── -->
                <TextBlock Style="{StaticResource SectionHeader}">MAILBOX</TextBlock>
                <Grid Margin="0,4,0,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="14"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <StackPanel Grid.Column="0">
                        <TextBlock Style="{StaticResource RequiredLabel}">MailBox Name *</TextBlock>
                        <TextBox x:Name="txtMailBoxName"/>
                    </StackPanel>
                    <StackPanel Grid.Column="2">
                        <TextBlock Style="{StaticResource RequiredLabel}">Folder Name *</TextBlock>
                        <TextBox x:Name="txtMailBoxFolderName" Text="Inbox"
                                 ToolTip="Folder name, or slash-delimited path for nested folders e.g. Inbox/SubFolder/Nested"/>
                        <CheckBox x:Name="chkSearchSubFolders" Content="Search Sub Folders"
                                  Margin="0,4,0,0" VerticalAlignment="Center"
                                  ToolTip="Recursively search all sub folders of the specified folder."/>
                    </StackPanel>
                </Grid>

                <!-- ── FILTERS ── -->
                <TextBlock Style="{StaticResource SectionHeader}">FILTERS</TextBlock>

                <!-- Subject + Sender -->
                <Grid Margin="0,4,0,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="14"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <StackPanel Grid.Column="0">
                        <TextBlock Style="{StaticResource Label}">Subject</TextBlock>
                        <TextBox x:Name="txtFilterSubject" Height="38" AcceptsReturn="True"
                                 VerticalScrollBarVisibility="Auto" TextWrapping="Wrap"
                                 VerticalContentAlignment="Top"
                                 ToolTip="One value per line or comma-separated. OR logic. Leave blank to match all."/>
                    </StackPanel>
                    <StackPanel Grid.Column="2">
                        <TextBlock Style="{StaticResource Label}">Sender</TextBlock>
                        <TextBox x:Name="txtFilterSender" Height="38" AcceptsReturn="True"
                                 VerticalScrollBarVisibility="Auto" TextWrapping="Wrap"
                                 VerticalContentAlignment="Top"
                                 ToolTip="Partial match on display name or email. One value per line or comma-separated. OR logic."/>
                    </StackPanel>
                </Grid>

                <!-- To + CC -->
                <Grid Margin="0,0,0,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="14"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <StackPanel Grid.Column="0">
                        <TextBlock Style="{StaticResource Label}">To</TextBlock>
                        <TextBox x:Name="txtFilterTo" Height="38" AcceptsReturn="True"
                                 VerticalScrollBarVisibility="Auto" TextWrapping="Wrap"
                                 VerticalContentAlignment="Top"
                                 ToolTip="Partial match on To field. One value per line or comma-separated. OR logic."/>
                    </StackPanel>
                    <StackPanel Grid.Column="2">
                        <TextBlock Style="{StaticResource Label}">CC</TextBlock>
                        <TextBox x:Name="txtFilterCC" Height="38" AcceptsReturn="True"
                                 VerticalScrollBarVisibility="Auto" TextWrapping="Wrap"
                                 VerticalContentAlignment="Top"
                                 ToolTip="Partial match on CC field. One value per line or comma-separated. OR logic."/>
                    </StackPanel>
                </Grid>

                <!-- To/CC OR mode -->
                <CheckBox x:Name="chkToCcOr" Content="Use To OR CC (single list below)" Margin="0,4,0,0"/>
                <TextBox x:Name="txtFilterToCcOr" Height="38" AcceptsReturn="True" 
                            VerticalScrollBarVisibility="Auto" TextWrapping="Wrap" VerticalContentAlignment="Top" 
                            ToolTip="One value per line or comma-separated. Matches in To OR CC email addresses."/>


                <!-- BCC collapsed -->
                <Expander Margin="0,4,0,0" IsExpanded="False">
                    <Expander.Header>
                        <TextBlock FontSize="11" FontWeight="SemiBold" Foreground="#888888">BCC Filter</TextBlock>
                    </Expander.Header>
                    <TextBox x:Name="txtFilterBCC" Height="38" AcceptsReturn="True" Margin="0,4,0,0"
                             VerticalScrollBarVisibility="Auto" TextWrapping="Wrap"
                             VerticalContentAlignment="Top"
                             ToolTip="Partial match on BCC field. Only populated on sent items. One value per line or comma-separated. OR logic."/>
                </Expander>

                <!-- ── DATE FILTERS + OUTPUT OPTIONS on one row ── -->
                <TextBlock Style="{StaticResource SectionHeader}">DATE FILTERS &amp; OUTPUT OPTIONS</TextBlock>
                <Grid Margin="0,4,0,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="14"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="14"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="14"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="14"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <StackPanel Grid.Column="0">
                        <TextBlock Style="{StaticResource Label}">Days Back</TextBlock>
                        <TextBox x:Name="txtDaysBack" ToolTip="Integer. Overridden by explicit From/To dates."/>
                    </StackPanel>
                    <StackPanel Grid.Column="2">
                        <TextBlock Style="{StaticResource Label}">From Date</TextBlock>
                        <TextBox x:Name="txtFromDate" ToolTip="Accepted formats: yyyy-MM-dd, MM/dd/yyyy, dd MMM yyyy, etc."/>
                    </StackPanel>
                    <StackPanel Grid.Column="4">
                        <TextBlock Style="{StaticResource Label}">To Date</TextBlock>
                        <TextBox x:Name="txtToDate" ToolTip="Accepted formats: yyyy-MM-dd, MM/dd/yyyy, dd MMM yyyy, etc."/>
                    </StackPanel>
                    <StackPanel Grid.Column="6">
                        <TextBlock Style="{StaticResource Label}">Create Folder</TextBlock>
                        <ComboBox x:Name="cmbCreateFolderPreference"
                                  ToolTip="Which folder to create if none in Save To Folders exist.">
                            <ComboBoxItem>First</ComboBoxItem>
                            <ComboBoxItem>Last</ComboBoxItem>
                        </ComboBox>
                    </StackPanel>
                    <StackPanel Grid.Column="8">
                        <TextBlock Style="{StaticResource Label}">File Collision</TextBlock>
                        <ComboBox x:Name="cmbFileCollisionAction"
                                  ToolTip="Action when a file with the same name already exists.">
                            <ComboBoxItem>Suffix</ComboBoxItem>
                            <ComboBoxItem>Overwrite</ComboBoxItem>
                            <ComboBoxItem>Skip</ComboBoxItem>
                            <ComboBoxItem>Error</ComboBoxItem>
                        </ComboBox>
                    </StackPanel>
                </Grid>

                <!-- ── OUTPUT + LOGGING side-by-side ── -->
                <Grid Margin="0,6,0,4">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="20"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <!-- OUTPUT: Save To Folders with Browse button -->
                    <StackPanel Grid.Column="0">
                        <TextBlock Style="{StaticResource SectionHeader}">OUTPUT</TextBlock>
                        <TextBlock Style="{StaticResource RequiredLabel}">Save To Folders *</TextBlock>
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="6"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <TextBox x:Name="txtSaveToFolders" Grid.Column="0"
                                     Height="44" AcceptsReturn="True"
                                     VerticalScrollBarVisibility="Auto"
                                     TextWrapping="Wrap" VerticalContentAlignment="Top"
                                     ToolTip="One path per line or comma-separated. First existing folder is used."/>
                            <Button x:Name="btnBrowseSave" Grid.Column="2" Content="Browse"
                                    Style="{StaticResource BrowseButton}" Height="44"
                                    ToolTip="Browse for a folder to append to the list."/>
                        </Grid>
                    </StackPanel>

                    <!-- LOGGING: Log Path with Browse button -->
                    <StackPanel Grid.Column="2">
                        <TextBlock Style="{StaticResource SectionHeader}">LOGGING</TextBlock>
                        <TextBlock Style="{StaticResource RequiredLabel}">Log Path * (or enable No Log)</TextBlock>
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="6"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <TextBox x:Name="txtLogPath" Grid.Column="0"
                                     ToolTip="Folder for the CSV log file. Defaults to Save To Folder if blank."/>
                            <Button x:Name="btnBrowseLog" Grid.Column="2" Content="Browse"
                                    Style="{StaticResource BrowseButton}"
                                    ToolTip="Browse for the log output folder."/>
                        </Grid>
                        <StackPanel Orientation="Horizontal" Margin="0,6,0,0">
                            <CheckBox x:Name="chkSkipAlreadyDownloaded"
                                      Content="Skip Already Downloaded"
                                      IsChecked="True" Margin="0,0,20,0"
                                      VerticalAlignment="Center"/>
                            <CheckBox x:Name="chkNoLog" Content="No Log"
                                      VerticalAlignment="Center"/>
                        </StackPanel>
                    </StackPanel>
                </Grid>

            </StackPanel>
        </ScrollViewer>


        <!-- Output log panel -->
        <Border Grid.Column="1" Grid.Row="1"
                Background="#1E1E1E" BorderBrush="#444444" BorderThickness="0,1,0,0">
          <DockPanel>
            <DockPanel DockPanel.Dock="Top" Background="#2D2D2D" LastChildFill="False">
              <TextBlock Text="OUTPUT" Foreground="#888888"
                         FontSize="11" FontWeight="Bold"
                         Padding="12,6" DockPanel.Dock="Left" VerticalAlignment="Center"/>
              <!-- NEW: Copy button -->
              <Button x:Name="btnCopyLog" Content="Copy"
                      DockPanel.Dock="Right"
                      Background="#444444" Foreground="#CCCCCC"
                      Margin="4" Style="{StaticResource SidebarButton}"/>
              <Button x:Name="btnClearLog" Content="Clear" DockPanel.Dock="Right"
                      Background="#444444" Foreground="#CCCCCC"
                      Margin="4" Style="{StaticResource SidebarButton}"/>
            </DockPanel>

                <ScrollViewer x:Name="outputScroller" VerticalScrollBarVisibility="Auto">
                    <TextBlock x:Name="txtOutput" Foreground="#D4D4D4"
                               FontFamily="Consolas" FontSize="12"
                               Padding="12,8" TextWrapping="Wrap"/>
                </ScrollViewer>
            </DockPanel>
        </Border>

        <!-- Bottom bar -->
        <Border Grid.Column="1" Grid.Row="2"
                Background="#ECF0F1" BorderBrush="#DDDDDD" BorderThickness="0,1,0,0"
                Padding="16,0">
            <DockPanel VerticalAlignment="Center">
                <TextBlock x:Name="txtStatus" DockPanel.Dock="Left"
                           Foreground="#666666" VerticalAlignment="Center"/>
                <Button x:Name="btnCancel" Content="Cancel" DockPanel.Dock="Right"
                        Background="#E74C3C" Foreground="White"
                        FontWeight="Bold" FontSize="14" Padding="16,6"
                        BorderThickness="0" Cursor="Hand"
                        Visibility="Collapsed" Margin="0,0,8,0"/>

                <Button x:Name="btnRun" Content="Run" DockPanel.Dock="Right"
                        Width="120"
                        Background="#27AE60" Foreground="White"
                        FontWeight="Bold" FontSize="13" Padding="12,6"
                        BorderThickness="0" Cursor="Hand"/>

            </DockPanel>
        </Border>
    </Grid>
</Window>
"@

# ---------------------------------------------------------------------------
# Build window
# ---------------------------------------------------------------------------

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

$lstProfiles               = $window.FindName("lstProfiles")
$txtProfileName            = $window.FindName("txtProfileName")
$chkVerbose                = $window.FindName("chkVerbose")
$txtMailBoxName            = $window.FindName("txtMailBoxName")
$txtMailBoxFolderName      = $window.FindName("txtMailBoxFolderName")
$chkSearchSubFolders       = $window.FindName("chkSearchSubFolders")
$txtFilterSubject          = $window.FindName("txtFilterSubject")
$txtFilterSender           = $window.FindName("txtFilterSender")
$txtFilterTo               = $window.FindName("txtFilterTo")
$txtFilterCC               = $window.FindName("txtFilterCC")
$txtFilterToCcOr           = $window.FindName('txtFilterToCcOr')
$chkToCcOr                 = $window.FindName('chkToCcOr')
$txtFilterBCC              = $window.FindName("txtFilterBCC")
$txtSaveToFolders          = $window.FindName("txtSaveToFolders")
$btnBrowseSave             = $window.FindName("btnBrowseSave")
$txtLogPath                = $window.FindName("txtLogPath")
$btnBrowseLog              = $window.FindName("btnBrowseLog")
$txtDaysBack               = $window.FindName("txtDaysBack")
$txtFromDate               = $window.FindName("txtFromDate")
$txtToDate                 = $window.FindName("txtToDate")
$chkSkipAlreadyDownloaded  = $window.FindName("chkSkipAlreadyDownloaded")
$chkNoLog                  = $window.FindName("chkNoLog")
$cmbFileCollisionAction    = $window.FindName("cmbFileCollisionAction")
$cmbCreateFolderPreference = $window.FindName("cmbCreateFolderPreference")
$btnCheckUpdate            = $window.FindName("btnCheckUpdate")
$btnNew                    = $window.FindName("btnNew")
$btnSave                   = $window.FindName("btnSave")
$btnDelete                 = $window.FindName("btnDelete")
$btnRun                    = $window.FindName("btnRun")
$btnCancel                 = $window.FindName("btnCancel")
$btnClearLog               = $window.FindName("btnClearLog")
$btnCopyLog                = $window.FindName("btnCopyLog")
$txtOutput                 = $window.FindName("txtOutput")
$outputScroller            = $window.FindName("outputScroller")
$txtStatus                 = $window.FindName("txtStatus")

# ---------------------------------------------------------------------------
# State
# ---------------------------------------------------------------------------

$global:configs        = [System.Collections.Generic.List[PSCustomObject]]::new()
$global:runProc        = $null   # holds the running process for Cancel support
$global:runOutFile     = $null   # temp file path for worker output
$global:runLinePos     = 0       # lines consumed from runOutFile so far
$global:pollTimer      = $null   # DispatcherTimer tailing the output file
$global:cancelTimer    = $null   # DispatcherTimer for graceful cancel
$global:cancelDeadline = $null   # deadline for graceful cancel
$global:isCancelling   = $false  # prevents re-entrant cancel clicks

foreach ($item in @(Load-Configs)) { $global:configs.Add($item) }

function Refresh-ProfileList {
    $lstProfiles.Items.Clear()
    foreach ($c in $global:configs) { [void]$lstProfiles.Items.Add($c.ProfileName) }
}

function Load-ProfileToForm($p) {
    $txtProfileName.Text                = $p.ProfileName
    $chkVerbose.IsChecked               =[bool]$p.Verbose
    $txtMailBoxName.Text                = if ($p.MailBoxName) { $p.MailBoxName } else { Get-DefaultMailBoxName }
    $txtMailBoxFolderName.Text          = if ($p.MailBoxFolderName) { $p.MailBoxFolderName } else { "Inbox" }
    $chkSearchSubFolders.IsChecked      = if ($p.PSObject.Properties['SearchSubFolders'])  { [bool]$p.SearchSubFolders  } else { $false }
    $txtFilterSubject.Text              = if ($p.PSObject.Properties['FilterSubject'])      { $p.FilterSubject } else { "" }
    $txtFilterSender.Text               = if ($p.PSObject.Properties['FilterSender'])       { $p.FilterSender  } else { "" }
    $txtFilterTo.Text                   = if ($p.PSObject.Properties['FilterTo'])           { $p.FilterTo      } else { "" }
    $txtFilterCC.Text                   = if ($p.PSObject.Properties['FilterCC'])           { $p.FilterCC      } else { "" }
    $txtFilterToCcOr.Text               = if ($p.PSObject.Properties['FilterToCcOr'])  { $p.FilterToCcOr }  else { '' }
    $chkToCcOr.IsChecked                = if ($p.PSObject.Properties['UseToCcOr'])     { [bool]$p.UseToCcOr } else { $false }
    $txtFilterBCC.Text                  = if ($p.PSObject.Properties['FilterBCC'])          { $p.FilterBCC     } else { "" }
    $txtSaveToFolders.Text              = $p.SaveToFolders
    $txtLogPath.Text                    = $p.LogPath
    $txtDaysBack.Text                   = $p.DaysBack
    $txtFromDate.Text                   = $p.FromDate
    $txtToDate.Text                     = $p.ToDate
    $chkSkipAlreadyDownloaded.IsChecked = [bool]$p.SkipAlreadyDownloaded
    $chkNoLog.IsChecked                 = [bool]$p.NoLog

    Set-DateFieldState $txtFromDate $true
    Set-DateFieldState $txtToDate   $true

    foreach ($item in $cmbFileCollisionAction.Items) {
        if ($item.Content -eq $p.FileCollisionAction) { $cmbFileCollisionAction.SelectedItem = $item; break }
    }
    foreach ($item in $cmbCreateFolderPreference.Items) {
        if ($item.Content -eq $p.CreateFolderPreference) { $cmbCreateFolderPreference.SelectedItem = $item; break }
    }

    #do not currently want this functionality -- going to implement a search that supports all??
    # Disable/enable To/CC when OR mode toggled
    $txtFilterTo.IsEnabled = -not [bool]$chkToCcOr.IsChecked
    $txtFilterCC.IsEnabled = -not [bool]$chkToCcOr.IsChecked

}

function Get-FormAsProfile {
    $collisionVal  = if ($cmbFileCollisionAction.SelectedItem)    { $cmbFileCollisionAction.SelectedItem.Content }    else { "Suffix" }
    $folderPrefVal = if ($cmbCreateFolderPreference.SelectedItem) { $cmbCreateFolderPreference.SelectedItem.Content } else { "Last" }

    return [PSCustomObject]@{
        ProfileName            = $txtProfileName.Text.Trim()
        Verbose                = [bool]$chkVerbose.IsChecked
        MailBoxName            = $txtMailBoxName.Text.Trim()
        MailBoxFolderName      = $txtMailBoxFolderName.Text.Trim()
        SearchSubFolders       = [bool]$chkSearchSubFolders.IsChecked
        FilterSubject          = $txtFilterSubject.Text
        FilterSender           = $txtFilterSender.Text
        FilterTo               = $txtFilterTo.Text
        FilterCC               = $txtFilterCC.Text
        FilterToCcOr  = $txtFilterToCcOr.Text
        UseToCcOr     = [bool]$chkToCcOr.IsChecked
        FilterBCC              = $txtFilterBCC.Text
        SaveToFolders          = $txtSaveToFolders.Text
        LogPath                = $txtLogPath.Text.Trim()
        DaysBack               = $txtDaysBack.Text.Trim()
        FromDate               = $txtFromDate.Text.Trim()
        ToDate                 = $txtToDate.Text.Trim()
        SkipAlreadyDownloaded  = [bool]$chkSkipAlreadyDownloaded.IsChecked
        NoLog                  = [bool]$chkNoLog.IsChecked
        FileCollisionAction    = $collisionVal
        CreateFolderPreference = $folderPrefVal
        RelativeBase           = $PSScriptRoot
    }
}

function Validate-Form {
    $errors = @()
    if ([string]::IsNullOrWhiteSpace($txtMailBoxName.Text))       { $errors += "- MailBox Name is required." }
    if ([string]::IsNullOrWhiteSpace($txtMailBoxFolderName.Text)) { $errors += "- MailBox Folder Name is required." }
    if ([string]::IsNullOrWhiteSpace($txtSaveToFolders.Text))     { $errors += "- Save To Folders is required." }
    if (-not $chkNoLog.IsChecked -and
        [string]::IsNullOrWhiteSpace($txtLogPath.Text))           { $errors += "- Log Path is required (or enable No Log)." }
    if (-not (Test-DateString $txtFromDate.Text.Trim()))          { $errors += "- From Date is not a valid date." }
    if (-not (Test-DateString $txtToDate.Text.Trim()))            { $errors += "- To Date is not a valid date." }
    return $errors
}

function Append-Output($text) {
    $txtOutput.Text += "$text`n"
    $outputScroller.ScrollToBottom()
}

function Format-ArrayArg($raw) {
    $vals = Parse-FilterValues $raw
    if ($vals.Count -eq 0) { return $null }
    $inner = ($vals | ForEach-Object { "`"$_`"" }) -join ","
    return "@($inner)"
}

function Build-WorkerCommand($p) {
    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add("& `"$ScriptPath`"")
    $lines.Add("-MailBoxName `"$($p.MailBoxName)`"")
    $lines.Add("-MailBoxFolderName `"$($p.MailBoxFolderName)`"")

    if ($p.SearchSubFolders) { $lines.Add("-SearchSubFolders") }

    $subjectArg = Format-ArrayArg $p.FilterSubject
    $senderArg  = Format-ArrayArg $p.FilterSender
    $toArg      = Format-ArrayArg $p.FilterTo
    $ccArg      = Format-ArrayArg $p.FilterCC
    $bccArg     = Format-ArrayArg $p.FilterBCC

    if ($subjectArg) { $lines.Add("-FilterSubject $subjectArg") }
    if ($senderArg)  { $lines.Add("-FilterSender $senderArg") }
    if ($toArg)      { $lines.Add("-FilterTo $toArg") }
    if ($ccArg)      { $lines.Add("-FilterCC $ccArg") }
    if ($bccArg)     { $lines.Add("-FilterBCC $bccArg") }

    if ($p.SaveToFolders) {
        $folderArg = Format-ArrayArg $p.SaveToFolders
        if ($folderArg) { $lines.Add("-SaveToFolders $folderArg") }
    }

    if ($p.LogPath -and -not $p.NoLog) { $lines.Add("-LogPath `"$($p.LogPath)`"") }
    if ($p.DaysBack)  { $lines.Add("-DaysBack $($p.DaysBack)") }
    if ($p.FromDate)  { $lines.Add("-FromDate `"$($p.FromDate)`"") }
    if ($p.ToDate)    { $lines.Add("-ToDate `"$($p.ToDate)`"") }

    $skipVal = if ($p.SkipAlreadyDownloaded) { '1' } else { '0' }
    $lines.Add("-SkipAlreadyDownloaded $skipVal")

    if ($p.NoLog) { $lines.Add("-NoLog") }

    $lines.Add("-FileCollisionAction `"$($p.FileCollisionAction)`"")
    $lines.Add("-CreateFolderPreference `"$($p.CreateFolderPreference)`"")
    $lines.Add("-RelativeBase `"$($p.RelativeBase)`"")

    if($p.Verbose){$lines.Add("-Verbose")}

    return $lines -join " "
}

function Build-ScriptArgs($p, [string]$OutFile) {
    # Wraps the worker command so all output streams go to a temp file.
    # The child process is launched fully detached (no pipe handles back to ISE)
    # which prevents COM/STA message-pump calls from freezing the UI thread.
    $workerCmd = (Build-WorkerCommand $p) -replace '"', '\"'
    $cmd = "$workerCmd *>&1 | Out-File -FilePath `"$OutFile`" -Encoding utf8 -Append"
    return "-NoProfile -ExecutionPolicy Bypass -Command `"$cmd`""
}

# ---------------------------------------------------------------------------
# Initialise
# ---------------------------------------------------------------------------

$cmbFileCollisionAction.SelectedIndex    = 0   # Suffix
$cmbCreateFolderPreference.SelectedIndex = 1   # Last

Refresh-ProfileList

$lastRun = Load-LastRun
if ($lastRun) {
    Load-ProfileToForm $lastRun
    $txtStatus.Text = "Last run profile restored."
} elseif ($global:configs.Count -gt 0) {
    Load-ProfileToForm $global:configs[0]
    $lstProfiles.SelectedIndex = 0
} else {
    $blank = New-EmptyProfile
    $blank.MailBoxName = Get-DefaultMailBoxName
    Load-ProfileToForm $blank
}

# ---------------------------------------------------------------------------
# Startup version check
# ---------------------------------------------------------------------------

$window.Add_Loaded({
    $window.Dispatcher.InvokeAsync([action]{
        Invoke-UpdateCheck -ManualTrigger $false
    }, [System.Windows.Threading.DispatcherPriority]::Background) | Out-Null
})

# ---------------------------------------------------------------------------
# Date field validation
# ---------------------------------------------------------------------------

$txtFromDate.Add_LostFocus({
    $valid = Test-DateString $txtFromDate.Text.Trim()
    Set-DateFieldState $txtFromDate $valid
    $txtStatus.Text = if (-not $valid) { "From Date: invalid date format." } else { "" }
})

$txtToDate.Add_LostFocus({
    $valid = Test-DateString $txtToDate.Text.Trim()
    Set-DateFieldState $txtToDate $valid
    $txtStatus.Text = if (-not $valid) { "To Date: invalid date format." } else { "" }
})

$txtFromDate.Add_TextChanged({ Set-DateFieldState $txtFromDate $true })
$txtToDate.Add_TextChanged({   Set-DateFieldState $txtToDate   $true })

# ---------------------------------------------------------------------------
# Events
# ---------------------------------------------------------------------------

$btnCheckUpdate.Add_Click({ Invoke-UpdateCheck -ManualTrigger $true })

# Browse for Save To Folders — appends to existing paths (multi-value field)
$btnBrowseSave.Add_Click({
    $picked = Show-FolderPicker
    if ($picked) {
        $existing = $txtSaveToFolders.Text.Trim()
        $txtSaveToFolders.Text = if ($existing) { "$existing`n$picked" } else { $picked }
    }
})

# Browse for Log Path — replaces the value (single path field)
$btnBrowseLog.Add_Click({
    $picked = Show-FolderPicker -InitialPath $txtLogPath.Text.Trim()
    if ($picked) { $txtLogPath.Text = $picked }
})

$lstProfiles.Add_SelectionChanged({
    $idx = $lstProfiles.SelectedIndex
    if ($idx -ge 0 -and $idx -lt $global:configs.Count) {
        Load-ProfileToForm $global:configs[$idx]
    }
})

$btnNew.Add_Click({
    $blank = New-EmptyProfile
    $blank.MailBoxName = Get-DefaultMailBoxName
    $global:configs.Add($blank)
    Save-Configs $global:configs
    Refresh-ProfileList
    $lstProfiles.SelectedIndex = $global:configs.Count - 1
    Load-ProfileToForm $blank
    $txtProfileName.Focus()
})

$btnSave.Add_Click({
    $profile = Get-FormAsProfile
    if ([string]::IsNullOrWhiteSpace($profile.ProfileName)) {
        [System.Windows.MessageBox]::Show(
            "Please enter a Profile Name before saving.",
            "Validation",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning) | Out-Null
        return
    }

    $existingIdx = -1
    for ($i = 0; $i -lt $global:configs.Count; $i++) {
        if ($global:configs[$i].ProfileName -eq $profile.ProfileName) { $existingIdx = $i; break }
    }

    if ($existingIdx -ge 0) {
        $global:configs[$existingIdx] = $profile
    } elseif ($lstProfiles.SelectedIndex -ge 0) {
        $global:configs[$lstProfiles.SelectedIndex] = $profile
    } else {
        $global:configs.Add($profile)
    }

    Save-Configs $global:configs
    Refresh-ProfileList

    for ($i = 0; $i -lt $lstProfiles.Items.Count; $i++) {
        if ($lstProfiles.Items[$i] -eq $profile.ProfileName) { $lstProfiles.SelectedIndex = $i; break }
    }
    $txtStatus.Text = "Profile '$($profile.ProfileName)' saved."
})

$btnDelete.Add_Click({
    $idx = $lstProfiles.SelectedIndex
    if ($idx -lt 0) { return }
    $name   = $global:configs[$idx].ProfileName
    $choice = [System.Windows.MessageBox]::Show(
        "Delete profile '$name'?",
        "Confirm Delete",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question)
    if ($choice -eq [System.Windows.MessageBoxResult]::Yes) {
        $global:configs.RemoveAt($idx)
        Save-Configs $global:configs
        Refresh-ProfileList
        if ($global:configs.Count -gt 0) {
            $lstProfiles.SelectedIndex = [Math]::Min($idx, $global:configs.Count - 1)
            Load-ProfileToForm $global:configs[$lstProfiles.SelectedIndex]
        } else {
            Load-ProfileToForm (New-EmptyProfile)
        }
        $txtStatus.Text = "Profile '$name' deleted."
    }
})

$btnClearLog.Add_Click({ $txtOutput.Text = "" })


$btnCopyLog.Add_Click({
    try {
        [System.Windows.Clipboard]::SetText($txtOutput.Text)
        $txtStatus.Text = "Output copied to clipboard."
    } catch {
        [System.Windows.MessageBox]::Show(
            "Unable to copy output to clipboard.`n`nError: $_",
            "Copy Failed",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        ) | Out-Null
    }
})

$chkToCcOr.Add_Checked({ 
    $txtFilterTo.IsEnabled = $false
    $txtFilterCC.IsEnabled = $false 
})

$chkToCcOr.Add_Unchecked({
    $txtFilterTo.IsEnabled = $true
    $txtFilterCC.IsEnabled = $true 
})


$btnCancel.Add_Click({
    if ($null -eq $global:runProc -or $global:runProc.HasExited) { return }
    if ($global:isCancelling) { return }
    $global:isCancelling = $true

    Append-Output "[$([datetime]::Now.ToString('HH:mm:ss'))] Cancel requested - waiting for worker to finish current folder..."
    $txtStatus.Text = "Cancelling..."

    # Stop the poll timer before starting the cancel timer to avoid races
    if ($null -ne $global:pollTimer) {
        $global:pollTimer.Stop()
        $global:pollTimer = $null
    }

    $global:cancelDeadline = [datetime]::Now.AddSeconds(5)
    $global:cancelTimer    = New-Object System.Windows.Threading.DispatcherTimer
    $global:cancelTimer.Interval = [TimeSpan]::FromMilliseconds(500)

    $global:cancelTimer.Add_Tick({
        $forced = $false

        if ($global:runProc -ne $null -and -not $global:runProc.HasExited) {
            if ([datetime]::Now -lt $global:cancelDeadline) { return }  # still waiting
            # Deadline passed — force kill
            try { $global:runProc.Kill() } catch {}
            $forced = $true
        }

        # Worker has exited (cleanly or forced) — stop timer first, then clean up
        $global:cancelTimer.Stop()
        $global:cancelTimer  = $null
        $global:isCancelling = $false

        # Final output drain
        if ($null -ne $global:runOutFile -and (Test-Path $global:runOutFile)) {
            $allLines = @(Get-Content -Path $global:runOutFile -Encoding utf8 -ErrorAction SilentlyContinue)
            for ($i = $global:runLinePos; $i -lt $allLines.Count; $i++) { Append-Output $allLines[$i] }
            try { Remove-Item $global:runOutFile -Force -ErrorAction SilentlyContinue } catch {}
        }
        $global:runProc    = $null
        $global:runOutFile = $null

        $btnRun.IsEnabled     = $true
        $btnCancel.Visibility = [System.Windows.Visibility]::Collapsed

        if ($forced) {
            Append-Output "[$([datetime]::Now.ToString('HH:mm:ss'))] Worker force-killed after timeout."
            $txtStatus.Text = "Cancelled (forced)."
        } else {
            Append-Output "[$([datetime]::Now.ToString('HH:mm:ss'))] Worker exited cleanly."
            $txtStatus.Text = "Cancelled."
        }
    })

    $global:cancelTimer.Start()
})

$btnRun.Add_Click({
    $errors = Validate-Form
    if ($errors.Count -gt 0) {
        [System.Windows.MessageBox]::Show(
            "Please fix the following before running:`n`n" + ($errors -join "`n"),
            "Validation Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning) | Out-Null
        return
    }

    $profile = Get-FormAsProfile
    Save-LastRun $profile

    # Temp file receives all worker output (stdout + stderr merged).
    # Using a file instead of pipes means no handles are inherited by ISE,
    # which eliminates the COM/STA message-pump interference that freezes the UI.
    $global:runOutFile = Join-Path $env:TEMP ("MARU_run_{0}.txt" -f [datetime]::Now.ToString("yyyyMMdd_HHmmss"))
    "" | Out-File -FilePath $global:runOutFile -Encoding utf8   # create/truncate

    $argList = Build-ScriptArgs $profile $global:runOutFile

    $txtOutput.Text = ""
    Append-Output "[$([datetime]::Now.ToString('HH:mm:ss'))] Starting: $($profile.ProfileName)"
    Append-Output "----------------------------------------"
    $txtStatus.Text       = "Running..."
    $btnRun.IsEnabled     = $false
    $btnCancel.Visibility = [System.Windows.Visibility]::Visible

    write-verbose 'commandline: $arglist'

    try {
        $proc = Start-Process -FilePath "powershell.exe" `
                              -ArgumentList $argList `
                              -WindowStyle Hidden `
                              -PassThru
        $global:runProc = $proc
    } catch {
        Append-Output "[FATAL] $_"
        $txtStatus.Text       = "Run failed."
        $btnRun.IsEnabled     = $true
        $btnCancel.Visibility = [System.Windows.Visibility]::Collapsed
        return
    }

    $global:runLinePos = 0
    $global:pollTimer  = New-Object System.Windows.Threading.DispatcherTimer
    $global:pollTimer.Interval = [TimeSpan]::FromMilliseconds(200)

    $global:pollTimer.Add_Tick({
        # Guard: cancel handler may have stopped and nulled this timer
        if ($null -eq $global:pollTimer) { return }
        if ($null -eq $global:runOutFile) { $global:pollTimer.Stop(); $global:pollTimer = $null; return }

        if (Test-Path $global:runOutFile) {
            $allLines = @(Get-Content -Path $global:runOutFile -Encoding utf8 -ErrorAction SilentlyContinue)
            for ($i = $global:runLinePos; $i -lt $allLines.Count; $i++) {
                Append-Output $allLines[$i]
            }
            $global:runLinePos = $allLines.Count
        }

        if ($global:runProc -ne $null -and $global:runProc.HasExited) {
            $global:pollTimer.Stop()
            $global:pollTimer = $null

            if ($null -ne $global:runOutFile -and (Test-Path $global:runOutFile)) {
                $allLines = @(Get-Content -Path $global:runOutFile -Encoding utf8 -ErrorAction SilentlyContinue)
                for ($i = $global:runLinePos; $i -lt $allLines.Count; $i++) {
                    Append-Output $allLines[$i]
                }
            }

            $exitCode = $global:runProc.ExitCode
            Append-Output "----------------------------------------"
            Append-Output "[$([datetime]::Now.ToString('HH:mm:ss'))] Completed. Exit code: $exitCode"
            $txtStatus.Text       = if ($exitCode -eq 0) { "Run completed successfully." } else { "Run finished with errors (exit code $exitCode)." }
            $btnRun.IsEnabled     = $true
            $btnCancel.Visibility = [System.Windows.Visibility]::Collapsed

            try { Remove-Item $global:runOutFile -Force -ErrorAction SilentlyContinue } catch {}
            $global:runProc    = $null
            $global:runOutFile = $null
        }
    })

    $global:pollTimer.Start()
})

# ---------------------------------------------------------------------------
# Show
# ---------------------------------------------------------------------------

[void]$window.ShowDialog()