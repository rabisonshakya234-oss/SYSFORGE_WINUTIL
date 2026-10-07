#Requires -Version 5.1
<#
.SYNOPSIS
    SysForge - a Windows system management utility (WPF GUI), inspired by Chris Titus Tech's WinUtil.

.DESCRIPTION
    Tabs:  Install | Tweaks | Config | Updates | Win11 Creator | Win10 Creator
      * Install        - categorized app installer (WinGet or Chocolatey), search, filters, upgrade-all
      * Tweaks         - essential/advanced tweaks with Undo, live preference toggles, DNS changer
      * Config         - Windows features, common fixes, legacy control panels, OpenSSH
      * Updates        - Recommended / Default / Disabled Windows Update profiles
      * Win11/Win10 Creator - take an OFFICIAL Microsoft ISO, debloat it offline, rebuild a new ISO

.NOTES
    Save as SysForge.ps1 and run:  powershell -ExecutionPolicy Bypass -File .\SysForge.ps1
    The script self-elevates to Administrator. Windows PowerShell 5.1 (built into Windows 10/11).
    Create a restore point (Tweaks > "Restore Point - Create") before applying tweaks.
#>

# ============================================================================
#  SELF-ELEVATION
# ============================================================================
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin -or $PSVersionTable.PSEdition -eq 'Core' -or [Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    if (-not $PSCommandPath) { Write-Host 'Please save this script as a .ps1 file and run it from disk.'; return }
    Start-Process powershell.exe -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -STA -File `"$PSCommandPath`""
    return
}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

$script:Log        = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]'
$script:Jobs       = New-Object System.Collections.ArrayList
$script:Busy       = $false
$script:LogTargets = @()
$script:LastMgr    = 'winget'
$SvcFile           = "$env:ProgramData\SysForge\services.json"

# ============================================================================
#  XAML  (main window)
# ============================================================================
$mainXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="SysForge - Windows Utility" Height="780" Width="1200" MinHeight="640" MinWidth="1000"
        WindowStartupLocation="CenterScreen" UseLayoutRounding="True"
        Background="{DynamicResource BgBrush}" Foreground="{DynamicResource FgBrush}"
        FontFamily="Segoe UI" FontSize="13">
  <Window.Resources>
    <SolidColorBrush x:Key="BgBrush" Color="#F3F3F3"/>
    <SolidColorBrush x:Key="PanelBrush" Color="#FFFFFF"/>
    <SolidColorBrush x:Key="FgBrush" Color="#1B1B1B"/>
    <SolidColorBrush x:Key="MutedBrush" Color="#6B6B6B"/>
    <SolidColorBrush x:Key="BorderBrushC" Color="#CFCFCF"/>
    <SolidColorBrush x:Key="AccentBrush" Color="#0F6CBD"/>
    <SolidColorBrush x:Key="HoverBrush" Color="#E8EEF5"/>

    <Style TargetType="Button">
      <Setter Property="Background" Value="{DynamicResource PanelBrush}"/>
      <Setter Property="Foreground" Value="{DynamicResource FgBrush}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushC}"/>
      <Setter Property="Padding" Value="12,7"/>
      <Setter Property="Margin" Value="0,3"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="1" CornerRadius="5" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="Background" Value="{DynamicResource HoverBrush}"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.45"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="NavTab" TargetType="RadioButton">
      <Setter Property="Foreground" Value="{DynamicResource FgBrush}"/>
      <Setter Property="Margin" Value="0,0,6,0"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="RadioButton">
            <Border x:Name="bd" Background="{DynamicResource PanelBrush}" BorderBrush="{DynamicResource BorderBrushC}"
                    BorderThickness="1" CornerRadius="5" Padding="16,7">
              <ContentPresenter HorizontalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="Background" Value="{DynamicResource HoverBrush}"/>
              </Trigger>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="bd" Property="Background" Value="{DynamicResource AccentBrush}"/>
                <Setter Property="Foreground" Value="White"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="Card" TargetType="CheckBox">
      <Setter Property="Foreground" Value="{DynamicResource FgBrush}"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="CheckBox">
            <Border x:Name="bd" Background="{DynamicResource PanelBrush}" BorderBrush="{DynamicResource BorderBrushC}"
                    BorderThickness="1" CornerRadius="6" Padding="12,9">
              <DockPanel>
                <TextBlock x:Name="chk" DockPanel.Dock="Right" Text="&#x2713;" FontWeight="Bold"
                           Foreground="{DynamicResource AccentBrush}" Visibility="Hidden"/>
                <ContentPresenter VerticalAlignment="Center"/>
              </DockPanel>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="Background" Value="{DynamicResource HoverBrush}"/>
              </Trigger>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="bd" Property="BorderBrush" Value="{DynamicResource AccentBrush}"/>
                <Setter TargetName="chk" Property="Visibility" Value="Visible"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="Switch" TargetType="CheckBox">
      <Setter Property="Foreground" Value="{DynamicResource FgBrush}"/>
      <Setter Property="Margin" Value="0,5"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="CheckBox">
            <StackPanel Orientation="Horizontal" Background="Transparent">
              <Grid Width="38" Height="20" Margin="0,0,10,0">
                <Border x:Name="track" CornerRadius="10" Background="{DynamicResource BorderBrushC}"/>
                <Ellipse x:Name="thumb" Width="14" Height="14" Fill="White" HorizontalAlignment="Left" Margin="3,0,0,0"/>
              </Grid>
              <ContentPresenter VerticalAlignment="Center"/>
            </StackPanel>
            <ControlTemplate.Triggers>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="track" Property="Background" Value="{DynamicResource AccentBrush}"/>
                <Setter TargetName="thumb" Property="HorizontalAlignment" Value="Right"/>
                <Setter TargetName="thumb" Property="Margin" Value="0,0,3,0"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style TargetType="CheckBox"><Setter Property="Foreground" Value="{DynamicResource FgBrush}"/></Style>
    <Style TargetType="RadioButton"><Setter Property="Foreground" Value="{DynamicResource FgBrush}"/><Setter Property="Margin" Value="0,3"/></Style>
    <Style TargetType="Expander"><Setter Property="Foreground" Value="{DynamicResource FgBrush}"/></Style>
    <Style TargetType="TextBox">
      <Setter Property="Background" Value="{DynamicResource PanelBrush}"/>
      <Setter Property="Foreground" Value="{DynamicResource FgBrush}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushC}"/>
      <Setter Property="CaretBrush" Value="{DynamicResource FgBrush}"/>
      <Setter Property="Padding" Value="6,4"/>
    </Style>
    <Style TargetType="ComboBox"><Setter Property="Foreground" Value="Black"/><Setter Property="Padding" Value="6,4"/></Style>
  </Window.Resources>

  <Grid Margin="12,10,12,8">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <DockPanel Grid.Row="0" Margin="0,0,0,10">
      <Button x:Name="btnTheme" DockPanel.Dock="Right" Content="&#x2699;" FontFamily="Segoe UI Symbol" FontSize="18" ToolTip="Settings" Padding="10,2" MinWidth="40"/>
      <StackPanel Orientation="Horizontal">
        <TextBlock Text="SysForge" FontSize="20" FontWeight="Bold" VerticalAlignment="Center" Margin="2,0,22,0"/>
        <RadioButton x:Name="navInstall"   Style="{StaticResource NavTab}" GroupName="nav" Content="Install" IsChecked="True"/>
        <RadioButton x:Name="navTweaks"    Style="{StaticResource NavTab}" GroupName="nav" Content="Tweaks"/>
        <RadioButton x:Name="navConfig"    Style="{StaticResource NavTab}" GroupName="nav" Content="Config"/>
        <RadioButton x:Name="navUpdates"   Style="{StaticResource NavTab}" GroupName="nav" Content="Updates"/>
        <RadioButton x:Name="navCreator11" Style="{StaticResource NavTab}" GroupName="nav" Content="Win11 Creator"/>
        <RadioButton x:Name="navCreator10" Style="{StaticResource NavTab}" GroupName="nav" Content="Win10 Creator"/>
      </StackPanel>
    </DockPanel>

    <Grid Grid.Row="1" x:Name="contentHost">

      <!-- ================= INSTALL ================= -->
      <Grid x:Name="viewInstall">
        <Grid.ColumnDefinitions><ColumnDefinition Width="230"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
        <Border Grid.Column="0" Background="{DynamicResource PanelBrush}" BorderBrush="{DynamicResource BorderBrushC}"
                BorderThickness="1" CornerRadius="6" Padding="12" Margin="0,0,10,0">
          <ScrollViewer VerticalScrollBarVisibility="Auto">
            <StackPanel>
              <TextBlock Text="Actions" FontWeight="Bold" FontSize="15" Margin="0,0,0,4"/>
              <Button x:Name="btnInstall" Content="Install / Upgrade Applications" Background="{DynamicResource AccentBrush}" Foreground="White"/>
              <Button x:Name="btnUninstall" Content="Uninstall Applications"/>
              <Button x:Name="btnUpgradeAll" Content="Upgrade All Applications"/>
              <TextBlock Text="Package Manager" FontWeight="Bold" FontSize="15" Margin="0,14,0,4"/>
              <RadioButton x:Name="rbWinget" Content="WinGet" GroupName="pm" IsChecked="True"/>
              <RadioButton x:Name="rbChoco" Content="Chocolatey" GroupName="pm"/>
              <TextBlock Text="Selection" FontWeight="Bold" FontSize="15" Margin="0,14,0,4"/>
              <Button x:Name="btnClear" Content="Clear Selection"/>
              <Button x:Name="btnCollapse" Content="Collapse All Categories"/>
              <Button x:Name="btnExpand" Content="Expand All Categories"/>
              <Button x:Name="btnDetect" Content="Detect Installed Apps"/>
              <TextBlock x:Name="lblSelected" Text="Selected apps: 0" Margin="2,10,0,0" Foreground="{DynamicResource MutedBrush}"/>
            </StackPanel>
          </ScrollViewer>
        </Border>
        <Grid Grid.Column="1">
          <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
          <StackPanel Grid.Row="0" Margin="0,0,0,8">
            <TextBox x:Name="txtSearch" ToolTip="Search applications" Margin="0,0,0,8"/>
            <WrapPanel x:Name="filterBar"/>
          </StackPanel>
          <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
            <StackPanel x:Name="appHost" Margin="0,0,8,0"/>
          </ScrollViewer>
        </Grid>
      </Grid>

      <!-- ================= TWEAKS ================= -->
      <Grid x:Name="viewTweaks" Visibility="Collapsed">
        <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
        <WrapPanel Grid.Row="0" Margin="0,0,0,8">
          <TextBlock Text="Recommended selections:" VerticalAlignment="Center" Margin="0,0,10,0"/>
          <Button x:Name="btnPMin" Content="Minimal" Margin="0,0,6,0" Width="100"/>
          <Button x:Name="btnPStd" Content="Standard" Margin="0,0,6,0" Width="100"/>
          <Button x:Name="btnPAdv" Content="Advanced" Margin="0,0,6,0" Width="100"/>
          <Button x:Name="btnPGet" Content="Get Installed" Margin="0,0,6,0" Width="110" ToolTip="Tick the tweaks that are already applied on this PC"/>
          <Button x:Name="btnPClr" Content="Clear" Margin="0,0,6,0" Width="100"/>
        </WrapPanel>
        <Grid Grid.Row="1">
          <Grid.ColumnDefinitions><ColumnDefinition Width="1.1*"/><ColumnDefinition Width="10"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
          <Border Grid.Column="0" Background="{DynamicResource PanelBrush}" BorderBrush="{DynamicResource BorderBrushC}" BorderThickness="1" CornerRadius="6" Padding="12">
            <DockPanel>
              <StackPanel DockPanel.Dock="Bottom" Orientation="Horizontal" Margin="0,8,0,0">
                <TextBlock Text="DNS:" VerticalAlignment="Center" Margin="0,0,8,0"/>
                <ComboBox x:Name="cmbDns" Width="220"/>
                <Button x:Name="btnDns" Content="Apply DNS" Margin="8,0,0,0" Padding="12,5"/>
              </StackPanel>
              <ScrollViewer VerticalScrollBarVisibility="Auto"><StackPanel x:Name="tweakHost"/></ScrollViewer>
            </DockPanel>
          </Border>
          <Border Grid.Column="2" Background="{DynamicResource PanelBrush}" BorderBrush="{DynamicResource BorderBrushC}" BorderThickness="1" CornerRadius="6" Padding="12">
            <ScrollViewer VerticalScrollBarVisibility="Auto"><StackPanel x:Name="prefHost"/></ScrollViewer>
          </Border>
        </Grid>
        <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,8,0,0">
          <Button x:Name="btnRunTweaks" Content="Run Tweaks" Width="170" Background="{DynamicResource AccentBrush}" Foreground="White"/>
          <Button x:Name="btnUndoTweaks" Content="Undo Selected Tweaks" Width="190" Margin="8,3,0,3"/>
          <Button x:Name="btnRestartExplorer" Content="Restart Explorer" Width="150" Margin="8,3,0,3"/>
          <Button x:Name="btnAppRemoval" Content="App Removal" Width="150" Margin="8,3,0,3" ToolTip="Install or remove built-in Windows (AppX) apps"/>
        </StackPanel>
      </Grid>

      <!-- ================= APP REMOVAL ================= -->
      <Grid x:Name="viewAppRemoval" Visibility="Collapsed">
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock Grid.Row="0" Text="Selections:" Margin="2,0,0,6"/>
        <WrapPanel Grid.Row="1" Margin="0,0,0,10">
          <Button x:Name="btnAppDefault" Content="Default" Width="250" Margin="0,0,6,0"/>
          <Button x:Name="btnAppGet" Content="Get Installed" Width="250" Margin="0,0,6,0"/>
          <Button x:Name="btnAppAll" Content="Select All" Width="250" Margin="0,0,6,0"/>
          <Button x:Name="btnAppClear" Content="Clear Selection" Width="250" Margin="0,0,6,0"/>
        </WrapPanel>
        <Grid Grid.Row="2">
          <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="12"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
          <Border Grid.Column="0" Background="{DynamicResource PanelBrush}" BorderBrush="{DynamicResource BorderBrushC}" BorderThickness="1" CornerRadius="6" Padding="12">
            <ScrollViewer VerticalScrollBarVisibility="Auto"><StackPanel x:Name="appLeft"/></ScrollViewer>
          </Border>
          <Border Grid.Column="2" Background="{DynamicResource PanelBrush}" BorderBrush="{DynamicResource BorderBrushC}" BorderThickness="1" CornerRadius="6" Padding="12">
            <ScrollViewer VerticalScrollBarVisibility="Auto"><StackPanel x:Name="appRight"/></ScrollViewer>
          </Border>
        </Grid>
        <Border Grid.Row="3" Margin="0,10,0,0" Padding="12" CornerRadius="6" BorderThickness="1" BorderBrush="{DynamicResource BorderBrushC}" Background="{DynamicResource PanelBrush}">
          <StackPanel>
            <TextBlock Text="Note: Select the Windows AppX packages you wish to install or remove." TextWrapping="Wrap" Margin="0,1"/>
            <TextBlock Text="Install Selected registers a local manifest when available, then falls back to the Microsoft Store." TextWrapping="Wrap" Margin="0,1"/>
            <TextBlock Text="Remove Selected removes packages for the current user and all new user profiles." TextWrapping="Wrap" Margin="0,1"/>
          </StackPanel>
        </Border>
        <WrapPanel Grid.Row="4" Margin="0,8,0,0">
          <Button x:Name="btnAppBack" Content="Back to Tweaks" Width="250" Margin="0,0,6,0"/>
          <Button x:Name="btnAppInstall" Content="Install Selected" Width="250" Margin="0,0,6,0"/>
          <Button x:Name="btnAppRemove" Content="Remove Selected" Width="250" Margin="0,0,6,0"/>
        </WrapPanel>
      </Grid>

      <!-- ================= CONFIG ================= -->
      <Grid x:Name="viewConfig" Visibility="Collapsed">
        <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="10"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
        <Border Grid.Column="0" Background="{DynamicResource PanelBrush}" BorderBrush="{DynamicResource BorderBrushC}" BorderThickness="1" CornerRadius="6" Padding="14">
          <ScrollViewer VerticalScrollBarVisibility="Auto"><StackPanel x:Name="cfgLeft"/></ScrollViewer>
        </Border>
        <Border Grid.Column="2" Background="{DynamicResource PanelBrush}" BorderBrush="{DynamicResource BorderBrushC}" BorderThickness="1" CornerRadius="6" Padding="14">
          <ScrollViewer VerticalScrollBarVisibility="Auto"><StackPanel x:Name="cfgRight"/></ScrollViewer>
        </Border>
      </Grid>

      <!-- ================= UPDATES ================= -->
      <ScrollViewer x:Name="viewUpdates" Visibility="Collapsed" VerticalScrollBarVisibility="Auto">
        <StackPanel Margin="4,0,4,0">
          <TextBlock Text="Windows Update Profiles" FontSize="26" FontWeight="Bold"/>
          <TextBlock Text="Choose how Windows receives updates. Each profile replaces the Windows Update settings managed by SysForge."
                     Foreground="{DynamicResource MutedBrush}" Margin="0,4,0,16" TextWrapping="Wrap"/>
          <Grid>
            <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
            <Border Grid.Column="0" Margin="0,0,12,0" Padding="20" CornerRadius="6" BorderThickness="2" BorderBrush="{DynamicResource AccentBrush}" Background="{DynamicResource PanelBrush}">
              <DockPanel MinHeight="300">
                <Button x:Name="btnUpdRec" DockPanel.Dock="Bottom" Content="Apply Recommended" Margin="0,16,0,0"/>
                <StackPanel>
                  <TextBlock Text="Recommended" FontSize="22" FontWeight="Bold"/>
                  <TextBlock Text="Balanced security and stability" Margin="0,2,0,14"/>
                  <TextBlock Text="- Defers feature updates for 365 days" Margin="0,3"/>
                  <TextBlock Text="- Defers quality updates for 4 days" Margin="0,3"/>
                  <TextBlock Text="- Excludes drivers from quality updates" Margin="0,3"/>
                  <TextBlock Text="- Prevents automatic restarts while a user is signed in" TextWrapping="Wrap" Margin="0,3"/>
                  <TextBlock Text="Deferral policies apply to Windows Pro, Enterprise and Education editions." FontStyle="Italic" TextWrapping="Wrap" Margin="0,12,0,0" Foreground="{DynamicResource MutedBrush}"/>
                </StackPanel>
              </DockPanel>
            </Border>
            <Border Grid.Column="1" Margin="0,0,12,0" Padding="20" CornerRadius="6" BorderThickness="1" BorderBrush="{DynamicResource BorderBrushC}" Background="{DynamicResource PanelBrush}">
              <DockPanel MinHeight="300">
                <Button x:Name="btnUpdDef" DockPanel.Dock="Bottom" Content="Restore Defaults" Margin="0,16,0,0"/>
                <StackPanel>
                  <TextBlock Text="Windows Default" FontSize="22" FontWeight="Bold"/>
                  <TextBlock Text="Return control to Windows" Margin="0,2,0,14"/>
                  <TextBlock Text="- Removes update policies applied by SysForge" TextWrapping="Wrap" Margin="0,3"/>
                  <TextBlock Text="- Restores update service startup settings" TextWrapping="Wrap" Margin="0,3"/>
                  <TextBlock Text="- Re-enables update scheduled tasks" Margin="0,3"/>
                  <TextBlock Text="Use this to undo the Recommended or Disable profile." FontStyle="Italic" Margin="0,12,0,0" TextWrapping="Wrap" Foreground="{DynamicResource MutedBrush}"/>
                </StackPanel>
              </DockPanel>
            </Border>
            <Border Grid.Column="2" Padding="20" CornerRadius="6" BorderThickness="1" BorderBrush="{DynamicResource BorderBrushC}" Background="{DynamicResource PanelBrush}">
              <DockPanel MinHeight="300">
                <Button x:Name="btnUpdDis" DockPanel.Dock="Bottom" Content="Disable Updates" Foreground="#D13438" Margin="0,16,0,0"/>
                <StackPanel>
                  <TextBlock Text="Disable Updates" FontSize="22" FontWeight="Bold" Foreground="#D13438"/>
                  <TextBlock Text="Advanced use only" FontWeight="SemiBold" Foreground="#D13438" Margin="0,2,0,14"/>
                  <TextBlock Text="- Disables the automatic update policy" TextWrapping="Wrap" Margin="0,3"/>
                  <TextBlock Text="- Stops update services and scheduled tasks" TextWrapping="Wrap" Margin="0,3"/>
                  <TextBlock Text="- Clears downloaded update files" Margin="0,3"/>
                  <TextBlock Text="Security updates will not be installed while this profile is active." FontStyle="Italic" Margin="0,12,0,0" TextWrapping="Wrap" Foreground="#D13438"/>
                </StackPanel>
              </DockPanel>
            </Border>
          </Grid>
          <Border Margin="0,18,0,0" Padding="14" CornerRadius="6" BorderThickness="1" BorderBrush="{DynamicResource BorderBrushC}" Background="{DynamicResource PanelBrush}">
            <TextBlock HorizontalAlignment="Center" TextWrapping="Wrap"
                       Text="Changes apply system-wide. Restart Windows after switching profiles. Use Restore Defaults to undo SysForge update policies."/>
          </Border>
        </StackPanel>
      </ScrollViewer>

    </Grid>

    <DockPanel Grid.Row="2" Margin="0,8,0,0">
      <ProgressBar x:Name="progress" DockPanel.Dock="Right" Width="140" Height="8" IsIndeterminate="True" Visibility="Collapsed"/>
      <TextBlock x:Name="lblStatus" Text="Ready" Foreground="{DynamicResource MutedBrush}"/>
    </DockPanel>
    <Expander x:Name="expLog" Grid.Row="3" Header="Activity log" Margin="0,4,0,0">
      <TextBox x:Name="txtLog" Height="130" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" FontFamily="Consolas" FontSize="12"/>
    </Expander>
  </Grid>
</Window>
'@

# ---- Windows 10 / Windows 11 creator tab layout (one template, used for both) ----
$creatorXaml = @'
<Grid xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Visibility="Collapsed">
  <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
  <Grid Grid.Row="0">
    <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="18"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
    <StackPanel Grid.Column="0">
      <TextBlock Text="Step 1 - Select Windows @@V@@ ISO" FontWeight="Bold"/>
      <TextBlock TextWrapping="Wrap" Margin="0,6,0,0" Text="Browse to your locally saved Windows @@V@@ ISO file. Only official ISOs downloaded from Microsoft are supported."/>
      <TextBlock TextWrapping="Wrap" Margin="0,8,0,8"><Run FontWeight="Bold" FontStyle="Italic" Text="NOTE:"/><Run FontStyle="Italic" Text=" This is only meant for Fresh and New Windows installs."/></TextBlock>
      <DockPanel>
        <Button x:Name="btnBrowse" DockPanel.Dock="Right" Content="Browse" Margin="8,0,0,0" Padding="16,6"/>
        <TextBox x:Name="txtIso" IsReadOnly="True" Text="No ISO selected..." VerticalContentAlignment="Center"/>
      </DockPanel>
      <Border x:Name="step2" Visibility="Collapsed" Margin="0,14,0,0" Padding="12" CornerRadius="6" BorderThickness="1"
              BorderBrush="{DynamicResource BorderBrushC}" Background="{DynamicResource PanelBrush}">
        <StackPanel>
          <TextBlock Text="Step 2 - Choose edition and options" FontWeight="Bold" Margin="0,0,0,8"/>
          <ComboBox x:Name="cmbEdition" Margin="0,0,0,8"/>
          <CheckBox x:Name="chkApps" Content="Remove preinstalled bloatware apps" IsChecked="True" Margin="0,3"/>
          <CheckBox x:Name="chkTele" Content="Disable telemetry, ads and suggestions" IsChecked="True" Margin="0,3"/>
          <CheckBox x:Name="chkOobe" Content="Skip privacy screens / allow local account setup" IsChecked="True" Margin="0,3"/>
          <CheckBox x:Name="chkOneDrive" Content="Remove OneDrive setup" Margin="0,3"/>
          <Button x:Name="btnCreate" Content="Create Windows @@V@@ ISO" Margin="0,10,0,0" Background="{DynamicResource AccentBrush}" Foreground="White"/>
        </StackPanel>
      </Border>
    </StackPanel>
    <Border Grid.Column="2" BorderThickness="1" BorderBrush="{DynamicResource BorderBrushC}" CornerRadius="5" Padding="14" VerticalAlignment="Top">
      <StackPanel>
        <TextBlock Text="!!WARNING!! You must use an official Microsoft ISO" FontWeight="Bold" Foreground="#E8450C"/>
        <TextBlock TextWrapping="Wrap" Margin="0,8,0,0" Text="Download the Windows @@V@@ ISO directly from Microsoft.com. Third-party, pre-modified, or unofficial images are not supported and may produce broken results."/>
        <TextBlock Text="On the Microsoft download page, choose:" Margin="0,10,0,0"/>
        <TextBlock Text="- Edition : Windows @@V@@" Margin="14,4,0,0"/>
        <TextBlock Text="- Language : your preferred language" Margin="14,2,0,0"/>
        <TextBlock Text="- Architecture : 64-bit (x64)" Margin="14,2,0,0"/>
        <Button x:Name="btnOpenMs" Content="Open Microsoft Download Page" HorizontalAlignment="Left" Margin="0,10,0,0" Padding="14,6"/>
      </StackPanel>
    </Border>
  </Grid>
  <TextBlock Grid.Row="1" Text="Status Log" FontWeight="Bold" Margin="0,14,0,4"/>
  <TextBox x:Name="txtStatus" Grid.Row="2" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
           VerticalContentAlignment="Top" Text="Ready. Please select a Windows @@V@@ ISO to begin."/>
</Grid>
'@

# ============================================================================
#  LOAD WINDOW
# ============================================================================
$window = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader ([xml]$mainXaml)))
foreach ($n in 'btnTheme','navInstall','navTweaks','navConfig','navUpdates','navCreator11','navCreator10','viewInstall','viewTweaks','viewConfig','viewUpdates',
               'contentHost','btnInstall','btnUninstall','btnUpgradeAll','rbWinget','rbChoco','btnClear','btnCollapse','btnExpand','btnDetect','lblSelected',
               'txtSearch','filterBar','appHost','btnPMin','btnPStd','btnPAdv','btnPClr','cmbDns','btnDns','tweakHost','prefHost','btnRunTweaks','btnUndoTweaks',
               'btnRestartExplorer','cfgLeft','cfgRight','btnUpdRec','btnUpdDef','btnUpdDis','progress','lblStatus','expLog','txtLog',
               'btnPGet','btnAppRemoval','viewAppRemoval','btnAppDefault','btnAppGet','btnAppAll','btnAppClear','appLeft','appRight','btnAppBack','btnAppInstall','btnAppRemove') {
    Set-Variable -Name $n -Value $window.FindName($n) -Scope Script
}
$script:LogTargets = @($txtLog)

# ============================================================================
#  GENERAL HELPERS
# ============================================================================
function New-Brush([string]$hex) { $c = [Windows.Media.ColorConverter]::ConvertFromString($hex); ([Windows.Media.SolidColorBrush]::new([Windows.Media.Color]$c)).psobject.BaseObject }

$Themes = @{
    Light = @{ BgBrush='#F3F3F3'; PanelBrush='#FFFFFF'; FgBrush='#1B1B1B'; MutedBrush='#6B6B6B'; BorderBrushC='#CFCFCF'; AccentBrush='#0F6CBD'; HoverBrush='#E8EEF5' }
    Dark  = @{ BgBrush='#1E1E1E'; PanelBrush='#2A2A2A'; FgBrush='#F0F0F0'; MutedBrush='#A5A5A5'; BorderBrushC='#4A4A4A'; AccentBrush='#2B7FD4'; HoverBrush='#3A3A3A' }
}
function Set-Theme([string]$name) {
    foreach ($k in $Themes[$name].Keys) { $window.Resources[$k] = New-Brush $Themes[$name][$k] }
    $script:Theme = $name
}

function Msg([string]$text, [string]$buttons = 'OK', [string]$icon = 'Information') {
    [Windows.MessageBox]::Show($window, $text, 'SysForge', $buttons, $icon)
}
function Set-Status([string]$t) { $lblStatus.Text = $t }
function Pump { $window.Dispatcher.Invoke([Action]{}, [Windows.Threading.DispatcherPriority]::Background) }
function Write-UILog([string]$m) { $script:Log.Enqueue($m) }

function Flush-Log {
    $sb = New-Object Text.StringBuilder
    [string]$line = $null; $last = $null
    while ($script:Log.TryDequeue([ref]$line)) { [void]$sb.AppendLine($line); $last = $line }
    if ($sb.Length -gt 0) {
        foreach ($t in $script:LogTargets) { $t.AppendText($sb.ToString()); $t.ScrollToEnd() }
        if ($last) { $lblStatus.Text = $last }
    }
}

function Set-Reg($Path, $Name, $Value, $Type = 'DWord') {
    if (-not (Test-Path $Path)) { New-Item -Path $Path -Force | Out-Null }
    New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType $Type -Force | Out-Null
}
function Remove-RegValue($Path, $Name) { Remove-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue }
function Get-Reg($Path, $Name, $Default = $null) {
    try { (Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop).$Name } catch { $Default }
}

# ---- background job runner (keeps the UI responsive) ----
function Start-BG {
    param([scriptblock]$Script, [object[]]$ArgList = @(), [scriptblock]$OnDone = $null, [string]$Label = 'Working...', [object[]]$Targets = $null)
    if ($script:Busy) { Msg 'Another task is still running. Please wait for it to finish.' 'OK' 'Warning'; return }
    if ($Targets) { $script:LogTargets = $Targets } else { $script:LogTargets = @($txtLog) }
    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = 'STA'; $rs.ThreadOptions = 'ReuseThread'; $rs.Open()
    $rs.SessionStateProxy.SetVariable('Log', $script:Log)
    $ps = [powershell]::Create(); $ps.Runspace = $rs
    [void]$ps.AddScript($Script.ToString())
    foreach ($a in $ArgList) { [void]$ps.AddArgument($a) }
    $script:Busy = $true
    $progress.Visibility = 'Visible'; Set-Status $Label
    [void]$script:Jobs.Add(@{ PS = $ps; RS = $rs; H = $ps.BeginInvoke(); Done = $OnDone })
}

$timer = New-Object Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromMilliseconds(250)
$timer.Add_Tick({
    Flush-Log
    foreach ($j in @($script:Jobs)) {
        if ($j.H.IsCompleted) {
            $res = $null
            try { $res = $j.PS.EndInvoke($j.H) } catch { $script:Log.Enqueue("ERROR: $($_.Exception.Message)") }
            $j.PS.Dispose(); $j.RS.Dispose(); [void]$script:Jobs.Remove($j)
            $script:Busy = $false; $progress.Visibility = 'Collapsed'
            Flush-Log
            if ($j.Done) { try { & $j.Done $res } catch { } }
            $script:LogTargets = @($txtLog)
            Set-Status 'Ready'
        }
    }
})
$timer.Start()

# ============================================================================
#  INSTALL TAB
# ============================================================================
$AppData = @'
Browsers|Brave|Brave.Brave|brave
Browsers|Google Chrome|Google.Chrome|googlechrome
Browsers|Chromium|Hibbiki.Chromium|chromium
Browsers|Microsoft Edge|Microsoft.Edge|microsoft-edge
Browsers|Firefox|Mozilla.Firefox|firefox
Browsers|LibreWolf|LibreWolf.LibreWolf|librewolf
Browsers|Mullvad Browser|MullvadVPN.MullvadBrowser|mullvad-browser
Browsers|Tor Browser|TorProject.TorBrowser|tor-browser
Browsers|Vivaldi|Vivaldi.Vivaldi|vivaldi
Browsers|Waterfox|Waterfox.Waterfox|waterfox
Browsers|Zen Browser|Zen-Team.Zen-Browser|
Communications|Discord|Discord.Discord|discord
Communications|Element|Element.Element|element-desktop
Communications|Microsoft Teams|Microsoft.Teams|microsoft-teams
Communications|Proton Mail|Proton.ProtonMail|protonmail
Communications|Signal|OpenWhisperSystems.Signal|signal
Communications|Slack|SlackTechnologies.Slack|slack
Communications|Telegram|Telegram.TelegramDesktop|telegram
Communications|Thunderbird|Mozilla.Thunderbird|thunderbird
Communications|Zoom|Zoom.Zoom|zoom
Development|Visual Studio Code|Microsoft.VisualStudioCode|vscode
Development|Git|Git.Git|git
Development|GitHub Desktop|GitHub.GitHubDesktop|github-desktop
Development|Windows Terminal|Microsoft.WindowsTerminal|microsoft-windows-terminal
Development|PowerShell 7|Microsoft.PowerShell|powershell-core
Development|Python 3|Python.Python.3.12|python
Development|Node.js LTS|OpenJS.NodeJS.LTS|nodejs-lts
Development|Go|GoLang.Go|golang
Development|Rustup|Rustlang.Rustup|rustup.install
Development|Temurin JDK 21|EclipseAdoptium.Temurin.21.JDK|temurin21
Development|Docker Desktop|Docker.DockerDesktop|docker-desktop
Development|Notepad++|Notepad++.Notepad++|notepadplusplus
Development|JetBrains Toolbox|JetBrains.Toolbox|jetbrainstoolbox
Development|Postman|Postman.Postman|postman
Development|WinSCP|WinSCP.WinSCP|winscp
Development|PuTTY|PuTTY.PuTTY|putty
Development|.NET SDK 8.0|Microsoft.DotNet.SDK.8|dotnet-8.0-sdk
Development|PostgreSQL 17|PostgreSQL.PostgreSQL.17|postgresql
Development|Oracle SQL Developer|Oracle.SQLDeveloper|sqldeveloper
Development|MongoDB Compass|MongoDB.Compass.Full|mongodb-compass
Documents|LibreOffice|TheDocumentFoundation.LibreOffice|libreoffice-fresh
Documents|Adobe Acrobat Reader|Adobe.Acrobat.Reader.64-bit|adobereader
Documents|SumatraPDF|SumatraPDF.SumatraPDF|sumatrapdf
Documents|Obsidian|Obsidian.Obsidian|obsidian
Documents|Notion|Notion.Notion|notion
Documents|Calibre|calibre.calibre|calibre
Games|Steam|Valve.Steam|steam
Games|Epic Games Launcher|EpicGames.EpicGamesLauncher|epicgameslauncher
Games|GOG Galaxy|GOG.Galaxy|goggalaxy
Games|Heroic Games Launcher|HeroicGamesLauncher.HeroicGamesLauncher|heroic-games-launcher
Games|Playnite|Playnite.Playnite|playnite
Games|Prism Launcher|PrismLauncher.PrismLauncher|prismlauncher
Microsoft Tools|PowerToys|Microsoft.PowerToys|powertoys
Microsoft Tools|.NET Desktop Runtime 8|Microsoft.DotNet.DesktopRuntime.8|dotnet-8.0-desktopruntime
Microsoft Tools|Visual C++ Redist 2015-2022 x64|Microsoft.VCRedist.2015+.x64|vcredist140
Microsoft Tools|Visual Studio 2022 Community|Microsoft.VisualStudio.2022.Community|visualstudio2022community
Microsoft Tools|Process Explorer|Microsoft.Sysinternals.ProcessExplorer|procexp
Microsoft Tools|OneDrive|Microsoft.OneDrive|onedrive
Multimedia|VLC|VideoLAN.VLC|vlc
Multimedia|mpv|shinchiro.mpv|mpv
Multimedia|OBS Studio|OBSProject.OBSStudio|obs-studio
Multimedia|Audacity|Audacity.Audacity|audacity
Multimedia|HandBrake|HandBrake.HandBrake|handbrake
Multimedia|FFmpeg|Gyan.FFmpeg|ffmpeg
Multimedia|GIMP|GIMP.GIMP|gimp
Multimedia|Krita|KDE.Krita|krita
Multimedia|Inkscape|Inkscape.Inkscape|inkscape
Multimedia|Blender|BlenderFoundation.Blender|blender
Multimedia|Kdenlive|KDE.Kdenlive|kdenlive
Multimedia|ShareX|ShareX.ShareX|sharex
Multimedia|ImageGlass|DuongDieuPhap.ImageGlass|imageglass
Multimedia|Spotify|Spotify.Spotify|spotify
Productivity|Todoist|Doist.Todoist|
Productivity|Joplin|Joplin.Joplin|joplin
Productivity|Logseq|Logseq.Logseq|logseq
Productivity|Anki|Anki.Anki|anki
Productivity|Zotero|DigitalScholar.Zotero|zotero
Productivity|Flow Launcher|Flow-Launcher.Flow-Launcher|flow-launcher
Productivity|Ditto Clipboard|Ditto.Ditto|ditto
Productivity|Greenshot|Greenshot.Greenshot|greenshot
Utilities|7-Zip|7zip.7zip|7zip
Utilities|WinRAR|RARLab.WinRAR|winrar
Utilities|NanaZip|M2Team.NanaZip|nanazip
Utilities|Everything|voidtools.Everything|everything
Utilities|WinDirStat|WinDirStat.WinDirStat|windirstat
Utilities|TreeSize Free|JAMSoftware.TreeSize.Free|treesizefree
Utilities|Rufus|Rufus.Rufus|rufus
Utilities|Ventoy|Ventoy.Ventoy|ventoy
Utilities|Balena Etcher|Balena.Etcher|etcher
Utilities|Bitwarden|Bitwarden.Bitwarden|bitwarden
Utilities|KeePassXC|KeePassXCTeam.KeePassXC|keepassxc
Utilities|qBittorrent|qBittorrent.qBittorrent|qbittorrent
Utilities|AutoHotkey|AutoHotkey.AutoHotkey|autohotkey
Utilities|CPU-Z|CPUID.CPU-Z|cpu-z
Utilities|HWiNFO|REALiX.HWiNFO|hwinfo
Utilities|CrystalDiskInfo|CrystalDewWorld.CrystalDiskInfo|crystaldiskinfo
Utilities|Tailscale|Tailscale.Tailscale|tailscale
Utilities|AnyDesk|AnyDeskSoftwareGmbH.AnyDesk|anydesk
Utilities|TeamViewer|TeamViewer.TeamViewer|teamviewer
'@
$Apps = $AppData -split "`r?`n" | Where-Object { $_.Trim() } | ForEach-Object {
    $p = $_.Split('|'); [pscustomobject]@{ Cat = $p[0]; Name = $p[1]; Winget = $p[2]; Choco = $p[3] }
}

$script:Cards = New-Object System.Collections.ArrayList
$script:Sections = [ordered]@{}
$script:ActiveCat = 'All'
$script:FilterButtons = @()

function Update-Count {
    $n = @($script:Cards | Where-Object { $_.IsChecked }).Count
    $lblSelected.Text = "Selected apps: $n"
}

function Apply-Filter {
    $q = $txtSearch.Text.Trim()
    foreach ($e in $script:Sections.Values) {
        $any = $false
        foreach ($cb in $e.Content.Children) {
            $ok = ($q -eq '') -or ($cb.Content.ToString().IndexOf($q, [StringComparison]::OrdinalIgnoreCase) -ge 0)
            if ($ok) { $cb.Visibility = 'Visible'; $any = $true } else { $cb.Visibility = 'Collapsed' }
        }
        if ($any -and ($script:ActiveCat -eq 'All' -or $e.Header -eq $script:ActiveCat)) { $e.Visibility = 'Visible' } else { $e.Visibility = 'Collapsed' }
    }
    foreach ($b in $script:FilterButtons) {
        if ($b.Tag -eq $script:ActiveCat) { $b.SetResourceReference([Windows.Controls.Control]::BackgroundProperty, 'HoverBrush') }
        else { $b.SetResourceReference([Windows.Controls.Control]::BackgroundProperty, 'PanelBrush') }
    }
}

$cats = @($Apps | Select-Object -ExpandProperty Cat -Unique)
foreach ($c in $cats) {
    $wp = New-Object Windows.Controls.WrapPanel
    $exp = New-Object Windows.Controls.Expander
    $exp.Header = $c; $exp.IsExpanded = $true; $exp.FontSize = 15; $exp.FontWeight = 'SemiBold'
    $exp.Margin = '0,0,0,10'; $exp.Content = $wp
    $wp.Margin = '0,8,0,0'; $wp.Children.Clear()
    foreach ($a in ($Apps | Where-Object { $_.Cat -eq $c })) {
        $cb = New-Object Windows.Controls.CheckBox
        $cb.Style = $window.FindResource('Card'); $cb.Content = $a.Name; $cb.Tag = $a
        $cb.Width = 225; $cb.Margin = '0,0,8,8'; $cb.FontSize = 13; $cb.FontWeight = 'Normal'
        $cb.ToolTip = "WinGet: $($a.Winget)    Chocolatey: $($a.Choco)"
        $cb.Add_Click({ Update-Count })
        [void]$wp.Children.Add($cb); [void]$script:Cards.Add($cb)
    }
    $script:Sections[$c] = $exp
    [void]$appHost.Children.Add($exp)
}

foreach ($c in (@('All') + $cats)) {
    $b = New-Object Windows.Controls.Button
    $b.Content = $c; $b.Tag = $c; $b.Padding = '12,4'; $b.Margin = '0,0,6,6'
    $b.Add_Click({ param($s, $e) $script:ActiveCat = [string]$s.Tag; Apply-Filter })
    $script:FilterButtons += $b
    [void]$filterBar.Children.Add($b)
}
Apply-Filter
$txtSearch.Add_TextChanged({ Apply-Filter })
$btnClear.Add_Click({ foreach ($cb in $script:Cards) { $cb.IsChecked = $false }; Update-Count })
$btnCollapse.Add_Click({ foreach ($e in $script:Sections.Values) { $e.IsExpanded = $false } })
$btnExpand.Add_Click({ foreach ($e in $script:Sections.Values) { $e.IsExpanded = $true } })

$script:AppJob = {
    param($Mode, $Mgr, $IdStr)
    [Console]::OutputEncoding = [Text.Encoding]::UTF8
    function L($m) { $Log.Enqueue([string]$m) }
    function Run($exe, $a) {
        & $exe @a 2>&1 | ForEach-Object {
            $t = ([string]$_).Trim()
            if ($t -and $t -notmatch '^[\-\\|/]+$' -and $t -notmatch '[\u2588\u2592]') { L $t }
        }
    }
    $ids = @(); if ($IdStr) { $ids = $IdStr -split '\|' }

    if ($Mgr -eq 'choco') {
        if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
            L 'Chocolatey not found - installing it first...'
            try {
                Set-ExecutionPolicy Bypass -Scope Process -Force
                [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor 3072
                Invoke-Expression ((New-Object Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1')) 2>&1 | Out-Null
                $env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' + [Environment]::GetEnvironmentVariable('Path', 'User')
            } catch { L "Chocolatey install failed: $($_.Exception.Message)" }
        }
        if (-not (Get-Command choco -ErrorAction SilentlyContinue)) { L 'ERROR: Chocolatey is not available.'; return }
    } else {
        if (-not (Get-Command winget -ErrorAction SilentlyContinue)) { L 'ERROR: WinGet not found. Use Config > Fixes > WinGet - Reinstall.'; return }
    }

    switch ($Mode) {
        'detect' {
            if ($Mgr -eq 'choco') { $o = & choco list -l -r 2>&1 } else { $o = & winget list --accept-source-agreements --disable-interactivity 2>&1 }
            ($o | Out-String)
            return
        }
        'upgradeall' {
            L 'Upgrading all applications...'
            if ($Mgr -eq 'choco') { Run 'choco' @('upgrade', 'all', '-y', '--no-progress') }
            else { Run 'winget' @('upgrade', '--all', '--silent', '--accept-package-agreements', '--accept-source-agreements', '--disable-interactivity') }
        }
        'install' {
            foreach ($id in $ids) {
                L ">>> Installing $id"
                if ($Mgr -eq 'choco') { Run 'choco' @('install', $id, '-y', '--no-progress') }
                else { Run 'winget' @('install', '--id', $id, '-e', '--silent', '--accept-package-agreements', '--accept-source-agreements', '--disable-interactivity') }
            }
        }
        'uninstall' {
            foreach ($id in $ids) {
                L ">>> Uninstalling $id"
                if ($Mgr -eq 'choco') { Run 'choco' @('uninstall', $id, '-y') }
                else { Run 'winget' @('uninstall', '--id', $id, '-e', '--silent', '--disable-interactivity') }
            }
        }
    }
    L 'Finished.'
}

function Invoke-AppAction([string]$mode) {
    if ($script:Busy) { Msg 'Another task is still running.' 'OK' 'Warning'; return }
    $mgr = 'winget'; if ($rbChoco.IsChecked) { $mgr = 'choco' }
    $script:LastMgr = $mgr
    $idStr = ''
    if ($mode -in 'install', 'uninstall') {
        $sel = @($script:Cards | Where-Object { $_.IsChecked } | ForEach-Object { $_.Tag })
        if ($sel.Count -eq 0) { Msg 'Select at least one application first.'; return }
        $ids = @($sel | ForEach-Object { if ($mgr -eq 'choco') { $_.Choco } else { $_.Winget } } | Where-Object { $_ })
        if ($ids.Count -eq 0) { Msg "None of the selected apps are available for $mgr."; return }
        $idStr = $ids -join '|'
        if ($mode -eq 'uninstall' -and (Msg "Uninstall $($ids.Count) application(s)?" 'YesNo' 'Question') -ne 'Yes') { return }
    }
    $done = $null
    if ($mode -eq 'detect') {
        $done = {
            param($res)
            $txt = [string]($res -join "`n")
            foreach ($cb in $script:Cards) {
                $id = $cb.Tag.Winget; if ($script:LastMgr -eq 'choco') { $id = $cb.Tag.Choco }
                if ($id) { $cb.IsChecked = [regex]::IsMatch($txt, '(?im)(?<![\w.\-])' + [regex]::Escape($id) + '(?![\w.\-])') }
            }
            Update-Count
            Write-UILog 'Installed applications detected and selected.'
        }
    }
    Start-BG -Script $script:AppJob -ArgList @($mode, $mgr, $idStr) -OnDone $done -Label "Running $mode ($mgr)..."
}
$btnInstall.Add_Click({ Invoke-AppAction 'install' })
$btnUninstall.Add_Click({ Invoke-AppAction 'uninstall' })
$btnUpgradeAll.Add_Click({ Invoke-AppAction 'upgradeall' })
$btnDetect.Add_Click({ Invoke-AppAction 'detect' })

# ============================================================================
#  TWEAKS TAB
# ============================================================================
#  G: E = essential, A = advanced.  P: preset level (1 minimal, 2 standard, 3 advanced, 9 = never preselected)
$Tweaks = @(
[pscustomobject]@{ G='E'; P=1; Name='Activity History - Disable'; Tip='Stops Windows from collecting and syncing your activity timeline.'
  Apply  = { $k='HKLM:\SOFTWARE\Policies\Microsoft\Windows\System'; 'EnableActivityFeed','PublishUserActivities','UploadUserActivities' | ForEach-Object { Set-Reg $k $_ 0 } }
  Revert = { $k='HKLM:\SOFTWARE\Policies\Microsoft\Windows\System'; 'EnableActivityFeed','PublishUserActivities','UploadUserActivities' | ForEach-Object { Remove-RegValue $k $_ } } }
[pscustomobject]@{ G='E'; P=1; Name='ConsumerFeatures - Disable'; Tip='Stops Windows from silently installing suggested apps and promotions.'
  Apply  = { Set-Reg 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent' 'DisableWindowsConsumerFeatures' 1 }
  Revert = { Remove-RegValue 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent' 'DisableWindowsConsumerFeatures' } }
[pscustomobject]@{ G='E'; P=2; Name='Delivery Optimization - Disable'; Tip='Stops Windows from using your bandwidth to upload updates to other PCs.'
  Apply  = { Set-Reg 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeliveryOptimization' 'DODownloadMode' 0 }
  Revert = { Remove-RegValue 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeliveryOptimization' 'DODownloadMode' } }
[pscustomobject]@{ G='E'; P=2; Name='End Task With Right Click - Enable'; Tip='Adds an End Task entry to the taskbar right-click menu (Windows 11).'
  Apply  = { Set-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\TaskbarDeveloperSettings' 'TaskbarEndTask' 1 }
  Revert = { Set-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\TaskbarDeveloperSettings' 'TaskbarEndTask' 0 } }
[pscustomobject]@{ G='E'; P=3; Name='Hibernation - Disable'; Tip='Turns off hibernation and Fast Startup; frees disk space. Not recommended for laptops.'
  Apply  = { powercfg.exe /hibernate off | Out-Null }
  Revert = { powercfg.exe /hibernate on | Out-Null } }
[pscustomobject]@{ G='E'; P=2; Name='Location Tracking - Disable'; Tip='Denies apps access to the system location service.'
  Apply  = { Set-Reg 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location' 'Value' 'Deny' 'String' }
  Revert = { Set-Reg 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location' 'Value' 'Allow' 'String' } }
[pscustomobject]@{ G='E'; P=1; Name='Restore Point - Create'; Tip='Creates a System Restore checkpoint before you change anything.'
  Apply  = { Enable-ComputerRestore -Drive "$env:SystemDrive\" -ErrorAction SilentlyContinue
             Set-Reg 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore' 'SystemRestorePointCreationFrequency' 0
             Checkpoint-Computer -Description 'SysForge restore point' -RestorePointType MODIFY_SETTINGS }
  Revert = $null }
[pscustomobject]@{ G='E'; P=2; Name='Services - Set to Manual'; Tip='Sets rarely-needed services (Maps, Fax, Xbox helpers, ...) to start only on demand. Original values are backed up.'
  Apply  = {
    $names = 'MapsBroker','lfsvc','RetailDemo','Fax','WMPNetworkSvc','XblAuthManager','XblGameSave','XboxNetApiSvc','RemoteRegistry','WerSvc'
    $orig = @{}
    foreach ($n in $names) { $s = Get-Service $n -ErrorAction SilentlyContinue; if ($s) { $orig[$n] = [string]$s.StartType; if ($s.StartType -ne 'Disabled') { Set-Service $n -StartupType Manual -ErrorAction SilentlyContinue } } }
    if (-not (Test-Path $SvcFile)) { New-Item -ItemType Directory (Split-Path $SvcFile) -Force | Out-Null; $orig | ConvertTo-Json | Set-Content $SvcFile }
  }
  Revert = {
    if (Test-Path $SvcFile) { $j = Get-Content $SvcFile -Raw | ConvertFrom-Json; foreach ($p in $j.PSObject.Properties) { Set-Service $p.Name -StartupType $p.Value -ErrorAction SilentlyContinue } }
  } }
[pscustomobject]@{ G='E'; P=1; Name='Telemetry - Disable'; Tip='Sets diagnostic data to the minimum, disables the advertising ID and the DiagTrack service.'
  Apply  = {
    Set-Reg 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection' 'AllowTelemetry' 0
    Set-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo' 'Enabled' 0
    foreach ($s in 'DiagTrack','dmwappushservice') { Stop-Service $s -Force -ErrorAction SilentlyContinue; Set-Service $s -StartupType Disabled -ErrorAction SilentlyContinue }
  }
  Revert = {
    Remove-RegValue 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection' 'AllowTelemetry'
    Set-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo' 'Enabled' 1
    Set-Service DiagTrack -StartupType Automatic -ErrorAction SilentlyContinue; Set-Service dmwappushservice -StartupType Manual -ErrorAction SilentlyContinue
  } }
[pscustomobject]@{ G='E'; P=1; Name='Temporary Files - Remove'; Tip='Deletes files in your user and system TEMP folders.'
  Apply  = { Remove-Item "$env:TEMP\*" -Recurse -Force -ErrorAction SilentlyContinue; Remove-Item "$env:SystemRoot\Temp\*" -Recurse -Force -ErrorAction SilentlyContinue }
  Revert = $null }
[pscustomobject]@{ G='E'; P=2; Name='Widgets / News Feed - Disable'; Tip='Disables the Widgets board and the news and interests feed via policy.'
  Apply  = { Set-Reg 'HKLM:\SOFTWARE\Policies\Microsoft\Dsh' 'AllowNewsAndInterests' 0 }
  Revert = { Remove-RegValue 'HKLM:\SOFTWARE\Policies\Microsoft\Dsh' 'AllowNewsAndInterests' } }

[pscustomobject]@{ G='A'; P=3; Name='Background Apps - Disable'; Tip='Prevents Store apps from running in the background for the current user.'
  Apply  = { Set-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\BackgroundAccessApplications' 'GlobalUserDisabled' 1 }
  Revert = { Set-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\BackgroundAccessApplications' 'GlobalUserDisabled' 0 } }
[pscustomobject]@{ G='A'; P=9; Name='Date & Time - Set Time to UTC'; Tip='Treats the hardware clock as UTC. Only useful when dual-booting with Linux.'
  Apply  = { Set-Reg 'HKLM:\SYSTEM\CurrentControlSet\Control\TimeZoneInformation' 'RealTimeIsUniversal' 1 'QWord' }
  Revert = { Remove-RegValue 'HKLM:\SYSTEM\CurrentControlSet\Control\TimeZoneInformation' 'RealTimeIsUniversal' } }
[pscustomobject]@{ G='A'; P=9; Name='Disable Reserved Storage'; Tip='Frees about 7 GB that Windows reserves for updates.'
  Apply  = { DISM.exe /Online /Set-ReservedStorageState /State:Disabled | Out-Null }
  Revert = { DISM.exe /Online /Set-ReservedStorageState /State:Enabled | Out-Null } }
[pscustomobject]@{ G='A'; P=9; Name='IPv6 - Disable'; Tip='Disables IPv6 on all adapters. Can break some networks and apps.'
  Apply  = { Set-Reg 'HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters' 'DisabledComponents' 255 }
  Revert = { Remove-RegValue 'HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters' 'DisabledComponents' } }
[pscustomobject]@{ G='A'; P=9; Name='Microsoft OneDrive - Remove'; Tip='Uninstalls OneDrive. Make sure your files are not only stored there.'
  Apply  = {
    Stop-Process -Name OneDrive -Force -ErrorAction SilentlyContinue
    $p = "$env:SystemRoot\SysWOW64\OneDriveSetup.exe"; if (-not (Test-Path $p)) { $p = "$env:SystemRoot\System32\OneDriveSetup.exe" }
    if (Test-Path $p) { Start-Process $p '/uninstall' -Wait }
  }
  Revert = {
    $p = "$env:SystemRoot\SysWOW64\OneDriveSetup.exe"; if (-not (Test-Path $p)) { $p = "$env:SystemRoot\System32\OneDriveSetup.exe" }
    if (Test-Path $p) { Start-Process $p '/install' -Wait }
  } }
[pscustomobject]@{ G='A'; P=3; Name='Razer Software Auto-Install - Disable'; Tip='Stops Windows from auto-installing Razer software when a Razer device is plugged in.'
  Apply  = { Set-Reg 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Device Installer' 'DisableCoInstallers' 1 }
  Revert = { Set-Reg 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Device Installer' 'DisableCoInstallers' 0 } }
[pscustomobject]@{ G='A'; P=3; Name='Right-Click Menu Previous Layout - Enable'; Tip='Restores the classic full context menu (Windows 11).'
  Apply  = { $k='HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32'; New-Item $k -Force | Out-Null; Set-ItemProperty $k '(default)' '' }
  Revert = { Remove-Item 'HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}' -Recurse -Force -ErrorAction SilentlyContinue } }
[pscustomobject]@{ G='A'; P=3; Name='Storage Sense - Disable'; Tip='Stops Windows from automatically deleting temporary files and recycle bin content.'
  Apply  = { Set-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\StorageSense\Parameters\StoragePolicy' '01' 0 }
  Revert = { Set-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\StorageSense\Parameters\StoragePolicy' '01' 1 } }
[pscustomobject]@{ G='A'; P=3; Name='Teredo - Disable'; Tip='Disables the legacy Teredo IPv6 tunneling interface.'
  Apply  = { netsh interface teredo set state disabled | Out-Null }
  Revert = { netsh interface teredo set state default | Out-Null } }
[pscustomobject]@{ G='A'; P=3; Name='Visual Effects - Set to Best Performance'; Tip='Turns off most animations and transparency effects.'
  Apply  = { Set-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects' 'VisualFXSetting' 2 }
  Revert = { Set-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects' 'VisualFXSetting' 0 } }
[pscustomobject]@{ G='A'; P=3; Name='Windows AI (Copilot / Recall) - Disable'; Tip='Applies policies that turn off Copilot and AI data analysis (Recall).'
  Apply  = { Set-Reg 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot' 'TurnOffWindowsCopilot' 1; Set-Reg 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsAI' 'DisableAIDataAnalysis' 1 }
  Revert = { Remove-RegValue 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot' 'TurnOffWindowsCopilot'; Remove-RegValue 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsAI' 'DisableAIDataAnalysis' } }
[pscustomobject]@{ G='A'; P=9; Name='Brave Browser - Debloat'; Tip='Applies Brave policies that disable Rewards, Wallet, VPN, AI Chat (Leo) and the stats ping.'
  Apply  = {
    $k = 'HKLM:\SOFTWARE\Policies\BraveSoftware\Brave'
    Set-Reg $k 'BraveRewardsDisabled' 1
    Set-Reg $k 'BraveWalletDisabled' 1
    Set-Reg $k 'BraveVPNDisabled' 1
    Set-Reg $k 'BraveAIChatEnabled' 0
    Set-Reg $k 'BraveStatsPingEnabled' 0
  }
  Revert = {
    $k = 'HKLM:\SOFTWARE\Policies\BraveSoftware\Brave'
    'BraveRewardsDisabled','BraveWalletDisabled','BraveVPNDisabled','BraveAIChatEnabled','BraveStatsPingEnabled' | ForEach-Object { Remove-RegValue $k $_ }
  } }
[pscustomobject]@{ G='A'; P=9; Name='Microsoft Edge - Debloat'; Tip='Applies Edge policies that disable shopping assistant, rewards, sidebar, startup boost, background mode, recommendations, promotions and first-run screens.'
  Apply  = {
    $k = 'HKLM:\SOFTWARE\Policies\Microsoft\Edge'
    $pol = @{
      HideFirstRunExperience = 1; EdgeShoppingAssistantEnabled = 0; ShowMicrosoftRewards = 0; PersonalizationReportingEnabled = 0
      StartupBoostEnabled = 0; BackgroundModeEnabled = 0; HubsSidebarEnabled = 0; EdgeFollowEnabled = 0
      ShowRecommendationsEnabled = 0; NewTabPageContentEnabled = 0; MicrosoftEdgeInsiderPromotionEnabled = 0
      DiagnosticData = 0; UserFeedbackAllowed = 0; ConfigureDoNotTrack = 1; PromotionalTabsEnabled = 0
      EdgeCollectionsEnabled = 0; CryptoWalletEnabled = 0; SpotlightExperiencesAndRecommendationsEnabled = 0
      WebWidgetAllowed = 0; EdgeWorkspacesEnabled = 0
    }
    foreach ($n in $pol.Keys) { Set-Reg $k $n $pol[$n] }
  }
  Revert = {
    $k = 'HKLM:\SOFTWARE\Policies\Microsoft\Edge'
    'HideFirstRunExperience','EdgeShoppingAssistantEnabled','ShowMicrosoftRewards','PersonalizationReportingEnabled','StartupBoostEnabled','BackgroundModeEnabled',
    'HubsSidebarEnabled','EdgeFollowEnabled','ShowRecommendationsEnabled','NewTabPageContentEnabled','MicrosoftEdgeInsiderPromotionEnabled','DiagnosticData',
    'UserFeedbackAllowed','ConfigureDoNotTrack','PromotionalTabsEnabled','EdgeCollectionsEnabled','CryptoWalletEnabled','SpotlightExperiencesAndRecommendationsEnabled',
    'WebWidgetAllowed','EdgeWorkspacesEnabled' | ForEach-Object { Remove-RegValue $k $_ }
  } }
[pscustomobject]@{ G='A'; P=9; Name='Microsoft Edge - Remove'; Tip='Uninstalls Microsoft Edge. WebView2 stays. Windows Update may reinstall it. Make sure you have another browser first. Undo reinstalls Edge via WinGet.'
  Apply  = {
    Stop-Process -Name msedge, MicrosoftEdgeUpdate -Force -ErrorAction SilentlyContinue
    Set-Reg 'HKLM:\SOFTWARE\Microsoft\EdgeUpdate' 'AllowUninstall' '' 'String'
    $base  = "${env:ProgramFiles(x86)}\Microsoft\Edge\Application"
    $setup = Get-ChildItem "$base\*\Installer\setup.exe" -ErrorAction SilentlyContinue | Sort-Object FullName -Descending | Select-Object -First 1
    if (-not $setup) { throw 'Edge installer not found (Edge may already be removed).' }
    Start-Process $setup.FullName '--uninstall --system-level --verbose-logging --force-uninstall' -Wait
    Remove-Item "$env:PUBLIC\Desktop\Microsoft Edge.lnk", "$env:USERPROFILE\Desktop\Microsoft Edge.lnk" -Force -ErrorAction SilentlyContinue
  }
  Revert = {
    Remove-RegValue 'HKLM:\SOFTWARE\Microsoft\EdgeUpdate' 'AllowUninstall'
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) { throw 'WinGet not found. Install Edge manually from microsoft.com/edge.' }
    & winget install -e --id Microsoft.Edge --silent --accept-package-agreements --accept-source-agreements --disable-interactivity | Out-Null
  } }
)

$script:TweakChecks = New-Object System.Collections.ArrayList
function New-Heading([string]$t, [string]$color = $null) {
    $tb = New-Object Windows.Controls.TextBlock
    $tb.Text = $t; $tb.FontSize = 16; $tb.FontWeight = 'Bold'; $tb.Margin = '0,8,0,6'
    if ($color) { $tb.Foreground = New-Brush $color }
    $tb
}
[void]$tweakHost.Children.Add((New-Heading 'Essential Tweaks'))
foreach ($t in ($Tweaks | Where-Object { $_.G -eq 'E' })) {
    $cb = New-Object Windows.Controls.CheckBox; $cb.Content = $t.Name; $cb.Tag = $t; $cb.ToolTip = $t.Tip; $cb.Margin = '4,3'
    [void]$tweakHost.Children.Add($cb); [void]$script:TweakChecks.Add($cb)
}
[void]$tweakHost.Children.Add((New-Heading 'Advanced Tweaks - CAUTION' '#D13438'))
foreach ($t in ($Tweaks | Where-Object { $_.G -eq 'A' })) {
    $cb = New-Object Windows.Controls.CheckBox; $cb.Content = $t.Name; $cb.Tag = $t; $cb.ToolTip = $t.Tip; $cb.Margin = '4,3'
    [void]$tweakHost.Children.Add($cb); [void]$script:TweakChecks.Add($cb)
}

function Set-Preset([int]$lvl) {
    foreach ($cb in $script:TweakChecks) { $cb.IsChecked = ($lvl -gt 0) -and ($cb.Tag.P -le $lvl) }
}
$btnPMin.Add_Click({ Set-Preset 1 }); $btnPStd.Add_Click({ Set-Preset 2 }); $btnPAdv.Add_Click({ Set-Preset 3 }); $btnPClr.Add_Click({ Set-Preset 0 })

# ---- Get Installed (Tweaks): tick the tweaks that are already applied on this PC ----
$script:TweakDetect = @{
    'Activity History - Disable'                = { (Get-Reg 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\System' 'EnableActivityFeed' 1) -eq 0 }
    'ConsumerFeatures - Disable'                = { (Get-Reg 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent' 'DisableWindowsConsumerFeatures' 0) -eq 1 }
    'Delivery Optimization - Disable'           = { (Get-Reg 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeliveryOptimization' 'DODownloadMode' 99) -eq 0 }
    'End Task With Right Click - Enable'        = { (Get-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\TaskbarDeveloperSettings' 'TaskbarEndTask' 0) -eq 1 }
    'Hibernation - Disable'                     = { (Get-Reg 'HKLM:\SYSTEM\CurrentControlSet\Control\Power' 'HibernateEnabled' 1) -eq 0 }
    'Location Tracking - Disable'               = { (Get-Reg 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location' 'Value' 'Allow') -eq 'Deny' }
    'Services - Set to Manual'                  = { Test-Path $SvcFile }
    'Telemetry - Disable'                       = { (Get-Reg 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection' 'AllowTelemetry' 99) -eq 0 }
    'Widgets / News Feed - Disable'             = { (Get-Reg 'HKLM:\SOFTWARE\Policies\Microsoft\Dsh' 'AllowNewsAndInterests' 99) -eq 0 }
    'Background Apps - Disable'                 = { (Get-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\BackgroundAccessApplications' 'GlobalUserDisabled' 0) -eq 1 }
    'Date & Time - Set Time to UTC'             = { (Get-Reg 'HKLM:\SYSTEM\CurrentControlSet\Control\TimeZoneInformation' 'RealTimeIsUniversal' 0) -eq 1 }
    'IPv6 - Disable'                            = { (Get-Reg 'HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters' 'DisabledComponents' 0) -eq 255 }
    'Razer Software Auto-Install - Disable'     = { (Get-Reg 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Device Installer' 'DisableCoInstallers' 0) -eq 1 }
    'Right-Click Menu Previous Layout - Enable' = { Test-Path 'HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32' }
    'Storage Sense - Disable'                   = { (Get-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\StorageSense\Parameters\StoragePolicy' '01' 1) -eq 0 }
    'Visual Effects - Set to Best Performance'  = { (Get-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects' 'VisualFXSetting' 0) -eq 2 }
    'Windows AI (Copilot / Recall) - Disable'   = { (Get-Reg 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot' 'TurnOffWindowsCopilot' 0) -eq 1 }
    'Brave Browser - Debloat'                   = { (Get-Reg 'HKLM:\SOFTWARE\Policies\BraveSoftware\Brave' 'BraveRewardsDisabled' 0) -eq 1 }
    'Microsoft Edge - Debloat'                  = { (Get-Reg 'HKLM:\SOFTWARE\Policies\Microsoft\Edge' 'HubsSidebarEnabled' 1) -eq 0 }
    'Microsoft Edge - Remove'                   = { -not (Test-Path "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe") }
}
$btnPGet.Add_Click({
    $found = 0
    foreach ($cb in $script:TweakChecks) {
        $d = $script:TweakDetect[[string]$cb.Tag.Name]
        $on = $false
        if ($d) { try { $on = [bool](& $d) } catch { $on = $false } }
        $cb.IsChecked = $on
        if ($on) { $found++ }
    }
    Set-Status "Get Installed: $found applied tweak(s) detected and selected."
    Write-UILog "Get Installed: $found applied tweak(s) detected and selected."
})

function Invoke-Tweaks([bool]$undo) {
    $sel = @($script:TweakChecks | Where-Object { $_.IsChecked })
    if ($sel.Count -eq 0) { Msg 'Select at least one tweak first.'; return }
    $verb = 'apply'; if ($undo) { $verb = 'undo' }
    $hasAdv = @($sel | Where-Object { $_.Tag.G -eq 'A' }).Count -gt 0
    $warn = "Do you want to $verb $($sel.Count) tweak(s)?"
    if ($hasAdv) { $warn += "`n`nAdvanced tweaks are selected. Make sure you understand them (hover for details) and have a restore point." }
    if ((Msg $warn 'YesNo' 'Question') -ne 'Yes') { return }
    $window.Cursor = [Windows.Input.Cursors]::Wait
    $expLog.IsExpanded = $true
    foreach ($cb in $sel) {
        $t = $cb.Tag; $sb = $t.Apply; if ($undo) { $sb = $t.Revert }
        if (-not $sb) { Write-UILog "Skipped (nothing to undo): $($t.Name)"; continue }
        Write-UILog "Working: $($t.Name)"; Flush-Log; Pump
        try { & $sb; Write-UILog "  OK" } catch { Write-UILog "  FAILED: $($_.Exception.Message)" }
        Flush-Log
    }
    $window.Cursor = $null
    Write-UILog 'Tweaks finished. Some changes need a restart or sign-out.'; Flush-Log
}
$btnRunTweaks.Add_Click({ Invoke-Tweaks $false })
$btnUndoTweaks.Add_Click({ Invoke-Tweaks $true })
$btnRestartExplorer.Add_Click({ Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue; Set-Status 'Explorer restarted.' })

# ---- DNS ----
$DnsMap = [ordered]@{
    'Default (DHCP)'          = $null
    'Cloudflare (1.1.1.1)'    = @('1.1.1.1', '1.0.0.1')
    'Google (8.8.8.8)'        = @('8.8.8.8', '8.8.4.4')
    'Quad9 (9.9.9.9)'         = @('9.9.9.9', '149.112.112.112')
    'AdGuard (94.140.14.14)'  = @('94.140.14.14', '94.140.15.15')
}
foreach ($k in $DnsMap.Keys) { [void]$cmbDns.Items.Add($k) }
$cmbDns.SelectedIndex = 0
$btnDns.Add_Click({
    $choice = [string]$cmbDns.SelectedItem; $addr = $DnsMap[$choice]
    try {
        Get-NetAdapter | Where-Object { $_.Status -eq 'Up' } | ForEach-Object {
            if ($addr) { Set-DnsClientServerAddress -InterfaceIndex $_.ifIndex -ServerAddresses $addr }
            else { Set-DnsClientServerAddress -InterfaceIndex $_.ifIndex -ResetServerAddresses }
        }
        Clear-DnsClientCache
        Set-Status "DNS set to: $choice"
    } catch { Msg "Could not set DNS: $($_.Exception.Message)" 'OK' 'Error' }
})

# ---- Preferences (instant toggles) ----
$Prefs = @(
@{ Name='Dark Theme for Windows'; Tip='Use the dark theme for apps and the system.'
   Get={ (Get-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' 'AppsUseLightTheme' 1) -eq 0 }
   Set={ param($on) $v = 1; if ($on) { $v = 0 }; $k = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize'; Set-Reg $k 'AppsUseLightTheme' $v; Set-Reg $k 'SystemUsesLightTheme' $v } },
@{ Name='File Explorer - Show File Extensions'; Tip='Always show file name extensions.'
   Get={ (Get-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' 'HideFileExt' 1) -eq 0 }
   Set={ param($on) $v = 1; if ($on) { $v = 0 }; Set-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' 'HideFileExt' $v } },
@{ Name='File Explorer - Show Hidden Files'; Tip='Show hidden files and folders.'
   Get={ (Get-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' 'Hidden' 2) -eq 1 }
   Set={ param($on) $v = 2; if ($on) { $v = 1 }; Set-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' 'Hidden' $v } },
@{ Name='Game Mode'; Tip='Let Windows prioritize resources for games.'
   Get={ (Get-Reg 'HKCU:\Software\Microsoft\GameBar' 'AutoGameModeEnabled' 1) -eq 1 }
   Set={ param($on) $v = 0; if ($on) { $v = 1 }; Set-Reg 'HKCU:\Software\Microsoft\GameBar' 'AutoGameModeEnabled' $v; Set-Reg 'HKCU:\Software\Microsoft\GameBar' 'AllowAutoGameMode' $v } },
@{ Name='Lock Screen - Disable'; Tip='Skip the lock screen and go straight to sign-in.'
   Get={ (Get-Reg 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization' 'NoLockScreen' 0) -eq 1 }
   Set={ param($on) $v = 0; if ($on) { $v = 1 }; Set-Reg 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization' 'NoLockScreen' $v } },
@{ Name='Enable Long Paths'; Tip='Allow file paths longer than 260 characters.'
   Get={ (Get-Reg 'HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem' 'LongPathsEnabled' 0) -eq 1 }
   Set={ param($on) $v = 0; if ($on) { $v = 1 }; Set-Reg 'HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem' 'LongPathsEnabled' $v } },
@{ Name='BSoD Verbose Mode'; Tip='Show technical details on the blue screen.'
   Get={ (Get-Reg 'HKLM:\SYSTEM\CurrentControlSet\Control\CrashControl' 'DisplayParameters' 0) -eq 1 }
   Set={ param($on) $v = 0; if ($on) { $v = 1 }; Set-Reg 'HKLM:\SYSTEM\CurrentControlSet\Control\CrashControl' 'DisplayParameters' $v } },
@{ Name='Num Lock on Startup'; Tip='Turn Num Lock on at the sign-in screen and after sign-in.'
   Get={ (Get-Reg 'HKCU:\Control Panel\Keyboard' 'InitialKeyboardIndicators' '0') -eq '2' }
   Set={ param($on) $v = '0'; if ($on) { $v = '2' }; Set-Reg 'HKCU:\Control Panel\Keyboard' 'InitialKeyboardIndicators' $v 'String'; Set-Reg 'Registry::HKEY_USERS\.DEFAULT\Control Panel\Keyboard' 'InitialKeyboardIndicators' $v 'String' } },
@{ Name='Sticky Keys Shortcut'; Tip='Pressing Shift five times opens the Sticky Keys prompt.'
   Get={ (Get-Reg 'HKCU:\Control Panel\Accessibility\StickyKeys' 'Flags' '510') -ne '506' }
   Set={ param($on) $v = '506'; if ($on) { $v = '510' }; Set-Reg 'HKCU:\Control Panel\Accessibility\StickyKeys' 'Flags' $v 'String' } },
@{ Name='Mouse Acceleration'; Tip='Enhance pointer precision.'
   Get={ (Get-Reg 'HKCU:\Control Panel\Mouse' 'MouseSpeed' '1') -ne '0' }
   Set={ param($on)
         $k = 'HKCU:\Control Panel\Mouse'
         if ($on) { Set-Reg $k 'MouseSpeed' '1' 'String'; Set-Reg $k 'MouseThreshold1' '6' 'String'; Set-Reg $k 'MouseThreshold2' '10' 'String' }
         else     { Set-Reg $k 'MouseSpeed' '0' 'String'; Set-Reg $k 'MouseThreshold1' '0' 'String'; Set-Reg $k 'MouseThreshold2' '0' 'String' } } },
@{ Name='Start Menu - Bing Web Search'; Tip='Include web results when searching from Start.'
   Get={ (Get-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Search' 'BingSearchEnabled' 1) -eq 1 }
   Set={ param($on) $v = 0; if ($on) { $v = 1 }; Set-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Search' 'BingSearchEnabled' $v } },
@{ Name='Taskbar - Search Icon'; Tip='Show the search icon/box on the taskbar.'
   Get={ (Get-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Search' 'SearchboxTaskbarMode' 1) -ne 0 }
   Set={ param($on) $v = 0; if ($on) { $v = 1 }; Set-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Search' 'SearchboxTaskbarMode' $v } },
@{ Name='Taskbar - Task View Icon'; Tip='Show the Task View button.'
   Get={ (Get-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' 'ShowTaskViewButton' 1) -eq 1 }
   Set={ param($on) $v = 0; if ($on) { $v = 1 }; Set-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' 'ShowTaskViewButton' $v } },
@{ Name='Taskbar - Centered Icons (Win 11)'; Tip='Center taskbar icons instead of left-aligning them.'
   Get={ (Get-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' 'TaskbarAl' 1) -eq 1 }
   Set={ param($on) $v = 0; if ($on) { $v = 1 }; Set-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' 'TaskbarAl' $v } },
@{ Name='Taskbar - Show Seconds in Clock'; Tip='Display seconds in the system tray clock.'
   Get={ (Get-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' 'ShowSecondsInSystemClock' 0) -eq 1 }
   Set={ param($on) $v = 0; if ($on) { $v = 1 }; Set-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' 'ShowSecondsInSystemClock' $v } }
)
[void]$prefHost.Children.Add((New-Heading 'Customize Preferences'))
$prefNote = New-Object Windows.Controls.TextBlock
$prefNote.Text = 'Toggles apply immediately. Use "Restart Explorer" if a change does not show up.'
$prefNote.FontStyle = 'Italic'; $prefNote.TextWrapping = 'Wrap'; $prefNote.Margin = '0,0,0,8'
$prefNote.SetResourceReference([Windows.Controls.TextBlock]::ForegroundProperty, 'MutedBrush')
[void]$prefHost.Children.Add($prefNote)
foreach ($p in $Prefs) {
    $cb = New-Object Windows.Controls.CheckBox
    $cb.Style = $window.FindResource('Switch'); $cb.Content = $p.Name; $cb.ToolTip = $p.Tip; $cb.Tag = $p
    try { $cb.IsChecked = [bool](& $p.Get) } catch { $cb.IsChecked = $false }
    $cb.Add_Click({
        param($s, $e)
        try { & $s.Tag.Set ([bool]$s.IsChecked); Set-Status "Applied: $($s.Tag.Name)" }
        catch { Msg "Could not change setting: $($_.Exception.Message)" 'OK' 'Error' }
    })
    [void]$prefHost.Children.Add($cb)
}

# ============================================================================
#  CONFIG TAB
# ============================================================================
function New-ActBtn([string]$text, [scriptblock]$click) {
    $b = New-Object Windows.Controls.Button
    $b.Content = $text; $b.Tag = $click; $b.Margin = '0,3'
    $b.Add_Click({ param($s, $e) & $s.Tag })
    $b
}

# ---- Power plans (added to the bottom of Tweaks > Customize Preferences) ----
$script:PlanGuid = @{
    Ultimate = 'e9a42b02-d5df-448d-aa00-03f14749eb61'
    High     = '8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c'
    Balanced = '381b4222-f694-41f0-9685-ff5bb260df2e'
}
function Get-PowerPlans {
    $plans = @()
    foreach ($line in @(powercfg.exe /list)) {
        if ($line -match '([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})\s+\((.+?)\)') {
            $plans += [pscustomobject]@{ Guid = $Matches[1].ToLower(); Name = $Matches[2] }
        }
    }
    $plans
}
function Enable-PowerPlan([string]$kind) {
    try {
        $base = $script:PlanGuid[$kind]
        $label = 'High Performance'; if ($kind -eq 'Ultimate') { $label = 'Ultimate Performance' }
        $target = $null
        $plans = @(Get-PowerPlans)
        if ($kind -eq 'High') { $hit = $plans | Where-Object { $_.Guid -eq $base } | Select-Object -First 1 }
        else { $hit = $plans | Where-Object { $_.Guid -eq $base -or $_.Name -eq $label } | Select-Object -First 1 }
        if ($hit) { $target = $hit.Guid }
        else {
            $out = (powercfg.exe -duplicatescheme $base | Out-String)
            if ($out -match '([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})') { $target = $Matches[1] }
        }
        if (-not $target) { throw "$label is not available on this system." }
        powercfg.exe /setactive $target
        if ($LASTEXITCODE -ne 0) { throw "powercfg could not activate $label." }
        Set-Status "$label power plan enabled."
        Write-UILog "$label power plan enabled."
    } catch { Msg "Could not enable the power plan: $($_.Exception.Message)" 'OK' 'Error' }
}
function Disable-PowerPlan([string]$kind) {
    try {
        $label = 'High Performance'; if ($kind -eq 'Ultimate') { $label = 'Ultimate Performance' }
        powercfg.exe /setactive $script:PlanGuid.Balanced
        if ($LASTEXITCODE -ne 0) { throw 'powercfg could not activate the Balanced plan.' }
        if ($kind -eq 'Ultimate') {
            foreach ($pl in @(Get-PowerPlans | Where-Object { $_.Guid -eq $script:PlanGuid.Ultimate -or $_.Name -eq $label })) { powercfg.exe /delete $pl.Guid | Out-Null }
        }
        Set-Status "$label power plan disabled (switched to Balanced)."
        Write-UILog "$label power plan disabled (switched to Balanced)."
    } catch { Msg "Could not disable the power plan: $($_.Exception.Message)" 'OK' 'Error' }
}
[void]$prefHost.Children.Add((New-Heading 'Performance Plans - NOT FOR LAPTOPS' '#D13438'))
[void]$prefHost.Children.Add((New-ActBtn 'Ultimate Performance - Enable'  { Enable-PowerPlan 'Ultimate' }))
[void]$prefHost.Children.Add((New-ActBtn 'Ultimate Performance - Disable' { Disable-PowerPlan 'Ultimate' }))
[void]$prefHost.Children.Add((New-ActBtn 'High Performance - Enable'      { Enable-PowerPlan 'High' }))
[void]$prefHost.Children.Add((New-ActBtn 'High Performance - Disable'     { Disable-PowerPlan 'High' }))

$Features = @(
    @{ N = '.NET Framework 3.5 (2.0 and 3.0) - Enable';  F = @('NetFx3') },
    @{ N = 'Hyper-V - Enable';                            F = @('Microsoft-Hyper-V-All') },
    @{ N = 'Legacy Media Components (WMP, DirectPlay)';   F = @('WindowsMediaPlayer', 'LegacyComponents', 'DirectPlay') },
    @{ N = 'Network File System (NFS) - Enable';          F = @('ServicesForNFS-ClientOnly', 'ClientForNFS-Infrastructure') },
    @{ N = 'Windows Sandbox - Enable';                    F = @('Containers-DisposableClientVM') },
    @{ N = 'Windows Subsystem for Linux (WSL) - Enable';  F = @('Microsoft-Windows-Subsystem-Linux', 'VirtualMachinePlatform') }
)
[void]$cfgLeft.Children.Add((New-Heading 'Features'))
$script:FeatChecks = @()
foreach ($f in $Features) {
    $cb = New-Object Windows.Controls.CheckBox; $cb.Content = $f.N; $cb.Tag = $f; $cb.Margin = '4,3'
    $script:FeatChecks += $cb; [void]$cfgLeft.Children.Add($cb)
}
[void]$cfgLeft.Children.Add((New-ActBtn 'Install Features' {
    $names = @($script:FeatChecks | Where-Object { $_.IsChecked } | ForEach-Object { $_.Tag.F })
    if ($names.Count -eq 0) { Msg 'Tick at least one feature first.'; return }
    Start-BG -Label 'Installing Windows features...' -ArgList @(($names -join '|')) -Script {
        param($FeatStr)
        function L($m) { $Log.Enqueue([string]$m) }
        foreach ($f in ($FeatStr -split '\|')) {
            L ">>> Enabling $f"
            try { $r = Enable-WindowsOptionalFeature -Online -FeatureName $f -All -NoRestart -ErrorAction Stop; L "  Done (restart needed: $($r.RestartNeeded))" }
            catch { L "  FAILED: $($_.Exception.Message)" }
        }
        L 'Finished. Restart Windows to complete feature installation.'
    }
}))

[void]$cfgLeft.Children.Add((New-Heading 'Fixes'))
$fixItems = @(
    @{ T = 'AutoLogon - Run'; A = { Start-Process netplwiz.exe } },
    @{ T = 'Network - Reset'; A = { Start-BG -Label 'Resetting network stack...' -Script {
            function L($m) { $Log.Enqueue([string]$m) }
            L 'Flushing DNS...';      ipconfig /flushdns | ForEach-Object { L $_ }
            L 'Resetting Winsock...'; netsh winsock reset | ForEach-Object { L $_ }
            L 'Resetting TCP/IP...';  netsh int ip reset | ForEach-Object { L $_ }
            L 'Done. Restart your PC to finish the network reset.' } } },
    @{ T = 'NTP Server - Enable'; A = {
            try {
                Set-Service w32time -StartupType Automatic; Start-Service w32time -ErrorAction SilentlyContinue
                w32tm /config /manualpeerlist:"time.windows.com pool.ntp.org" /syncfromflags:manual /reliable:yes /update | Out-Null
                w32tm /resync | Out-Null; Set-Status 'Time service configured and synchronized.'
            } catch { Msg "NTP setup failed: $($_.Exception.Message)" 'OK' 'Error' } } },
    @{ T = 'System Corruption Scan - Run'; A = { Start-Process cmd.exe -ArgumentList '/c sfc /scannow & DISM /Online /Cleanup-Image /RestoreHealth & pause' } },
    @{ T = 'Windows Update - Reset'; A = { Start-BG -Label 'Resetting Windows Update...' -Script {
            function L($m) { $Log.Enqueue([string]$m) }
            $svc = 'wuauserv', 'cryptSvc', 'bits', 'msiserver'
            foreach ($s in $svc) { L "Stopping $s"; Stop-Service $s -Force -ErrorAction SilentlyContinue }
            foreach ($d in "$env:SystemRoot\SoftwareDistribution", "$env:SystemRoot\System32\catroot2") {
                if (Test-Path "$d.bak") { Remove-Item "$d.bak" -Recurse -Force -ErrorAction SilentlyContinue }
                if (Test-Path $d) { Rename-Item $d "$(Split-Path $d -Leaf).bak" -ErrorAction SilentlyContinue; L "Renamed $d" }
            }
            foreach ($s in $svc) { L "Starting $s"; Start-Service $s -ErrorAction SilentlyContinue }
            L 'Windows Update components reset. Restart and check for updates again.' } } },
    @{ T = 'WinGet - Reinstall'; A = { Start-BG -Label 'Reinstalling WinGet...' -Script {
            function L($m) { $Log.Enqueue([string]$m) }
            try { Add-AppxPackage -RegisterByFamilyName -MainPackage Microsoft.DesktopAppInstaller_8wekyb3d8bbwe -ErrorAction Stop; L 'App Installer re-registered.' }
            catch {
                L 'Re-registering failed; downloading the latest App Installer...'
                try {
                    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor 3072
                    $f = Join-Path $env:TEMP 'AppInstaller.msixbundle'
                    Invoke-WebRequest 'https://aka.ms/getwinget' -OutFile $f -UseBasicParsing
                    Add-AppxPackage $f -ErrorAction Stop; L 'WinGet installed.'
                } catch { L "FAILED: $($_.Exception.Message)" }
            } } } }
)
foreach ($f in $fixItems) { [void]$cfgLeft.Children.Add((New-ActBtn $f.T $f.A)) }

[void]$cfgRight.Children.Add((New-Heading 'Legacy Windows Panels'))
$legacy = [ordered]@{
    'Computer Management'      = { Start-Process compmgmt.msc }
    'Control Panel'            = { Start-Process control.exe }
    'Mouse Properties'         = { Start-Process main.cpl }
    'Network Connections'      = { Start-Process ncpa.cpl }
    'Power Panel'              = { Start-Process powercfg.cpl }
    'Printer Panel'            = { Start-Process control.exe -ArgumentList 'printers' }
    'Programs and Features'    = { Start-Process appwiz.cpl }
    'Region'                   = { Start-Process intl.cpl }
    'Security and Maintenance' = { Start-Process control.exe -ArgumentList '/name Microsoft.ActionCenter' }
    'Sound Settings'           = { Start-Process mmsys.cpl }
    'System Properties'        = { Start-Process sysdm.cpl }
    'Time and Date'            = { Start-Process timedate.cpl }
    'Windows Defender Firewall'= { Start-Process firewall.cpl }
    'Windows Restore'          = { Start-Process rstrui.exe }
}
foreach ($k in $legacy.Keys) { [void]$cfgRight.Children.Add((New-ActBtn $k $legacy[$k])) }
[void]$cfgRight.Children.Add((New-Heading 'Remote Access'))
[void]$cfgRight.Children.Add((New-ActBtn 'OpenSSH Server - Enable' { Start-BG -Label 'Installing OpenSSH Server...' -Script {
    function L($m) { $Log.Enqueue([string]$m) }
    try {
        Add-WindowsCapability -Online -Name 'OpenSSH.Server~~~~0.0.1.0' -ErrorAction Stop | Out-Null
        Set-Service sshd -StartupType Automatic; Start-Service sshd
        if (-not (Get-NetFirewallRule -Name 'OpenSSH-Server-In-TCP' -ErrorAction SilentlyContinue)) {
            New-NetFirewallRule -Name 'OpenSSH-Server-In-TCP' -DisplayName 'OpenSSH Server (sshd)' -Enabled True -Direction Inbound -Protocol TCP -Action Allow -LocalPort 22 | Out-Null }
        L 'OpenSSH Server is installed and running on port 22.'
    } catch { L "FAILED: $($_.Exception.Message)" } } }))

# ============================================================================
#  APP REMOVAL (Tweaks > App Removal)  -  install / remove built-in AppX apps
# ============================================================================
#  Format:  Category | Display name | AppX package name | Default selection (1/0)
$AppxData = @'
Microsoft Apps|Feedback Hub|Microsoft.WindowsFeedbackHub|1
Microsoft Apps|Get Help|Microsoft.GetHelp|1
Microsoft Apps|Microsoft 365|Microsoft.MicrosoftOfficeHub|1
Microsoft Apps|Microsoft Teams|MSTeams|1
Microsoft Apps|Outlook for Windows|Microsoft.OutlookForWindows|1
Microsoft Ecosystem|Mobile Devices|MicrosoftWindows.CrossDevice|1
Microsoft Ecosystem|Phone Link|Microsoft.YourPhone|1
Utilities & Productivity|Calculator|Microsoft.WindowsCalculator|0
Utilities & Productivity|Camera|Microsoft.WindowsCamera|0
Utilities & Productivity|Clipchamp|Clipchamp.Clipchamp|1
Utilities & Productivity|Clock|Microsoft.WindowsAlarms|0
Utilities & Productivity|Media Player|Microsoft.ZuneMusic|0
Utilities & Productivity|Notepad|Microsoft.WindowsNotepad|0
Utilities & Productivity|Paint|Microsoft.Paint|0
Utilities & Productivity|Photos|Microsoft.Windows.Photos|0
Utilities & Productivity|Quick Assist|MicrosoftCorporationII.QuickAssist|0
Utilities & Productivity|Snipping Tool|Microsoft.ScreenSketch|0
Utilities & Productivity|Sound Recorder|Microsoft.WindowsSoundRecorder|0
Utilities & Productivity|Sticky Notes|Microsoft.MicrosoftStickyNotes|0
Utilities & Productivity|To Do|Microsoft.Todos|1
Bing & Web Services|Bing Search|Microsoft.BingSearch|1
Bing & Web Services|Copilot|Microsoft.Copilot|1
Bing & Web Services|News|Microsoft.BingNews|1
Bing & Web Services|Start Experiences App|MicrosoftWindows.Client.WebExperience|1
Bing & Web Services|Weather|Microsoft.BingWeather|1
Developer Tools|Dev Home|Microsoft.Windows.DevHome|1
Developer Tools|Power Automate|Microsoft.PowerAutomateDesktop|1
Xbox & Gaming|Solitaire Collection|Microsoft.MicrosoftSolitaireCollection|1
Xbox & Gaming|Xbox App|Microsoft.GamingApp|1
Xbox & Gaming|Xbox Game Bar|Microsoft.XboxGamingOverlay|1
Xbox & Gaming|Xbox Identity Provider|Microsoft.XboxIdentityProvider|1
Xbox & Gaming|Xbox Speech To Text Overlay|Microsoft.XboxSpeechToTextOverlay|1
Xbox & Gaming|Xbox TCUI|Microsoft.Xbox.TCUI|1
'@
$script:AppxItems = $AppxData -split "`r?`n" | Where-Object { $_.Trim() } | ForEach-Object {
    $p = $_.Split('|'); [pscustomobject]@{ Cat = $p[0]; Name = $p[1]; Pkg = $p[2]; Def = ($p[3] -eq '1') }
}
$script:AppxLeftCats = 'Microsoft Apps', 'Microsoft Ecosystem', 'Utilities & Productivity'
$script:AppxChecks = New-Object System.Collections.ArrayList
foreach ($c in @($script:AppxItems | Select-Object -ExpandProperty Cat -Unique)) {
    $host_ = $appRight; if ($script:AppxLeftCats -contains $c) { $host_ = $appLeft }
    $h = New-Heading $c; $h.FontFamily = New-Object Windows.Media.FontFamily 'Consolas'; $h.FontSize = 17
    [void]$host_.Children.Add($h)
    foreach ($a in ($script:AppxItems | Where-Object { $_.Cat -eq $c })) {
        $cb = New-Object Windows.Controls.CheckBox
        $cb.Content = $a.Name; $cb.Tag = $a; $cb.ToolTip = $a.Pkg; $cb.Margin = '36,3,0,3'
        [void]$host_.Children.Add($cb); [void]$script:AppxChecks.Add($cb)
    }
}

function Show-AppRemoval {
    foreach ($k in $script:Views.Keys) { $script:Views[$k].Visibility = 'Collapsed' }
    $viewAppRemoval.Visibility = 'Visible'
    $expLog.Visibility = 'Visible'
    $navTweaks.IsChecked = $false
}

$script:AppxJob = {
    param($Mode, $IdStr)
    function L($m) { $Log.Enqueue([string]$m) }
    $items = @(); if ($IdStr) { $items = $IdStr -split ';' }
    switch ($Mode) {
        'detect' {
            Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue | ForEach-Object { [string]$_.Name }
            return
        }
        'remove' {
            foreach ($it in $items) {
                $pkg, $disp = $it -split '=', 2
                L ">>> Removing $disp"
                try { Get-AppxPackage -Name $pkg -AllUsers -ErrorAction Stop | Remove-AppxPackage -AllUsers -ErrorAction Stop }
                catch {
                    try { Get-AppxPackage -Name $pkg -ErrorAction Stop | Remove-AppxPackage -ErrorAction Stop }
                    catch { L "  Could not remove for existing users: $($_.Exception.Message)" }
                }
                try {
                    Get-AppxProvisionedPackage -Online -ErrorAction Stop | Where-Object { $_.DisplayName -eq $pkg } | ForEach-Object {
                        Remove-AppxProvisionedPackage -Online -PackageName $_.PackageName -ErrorAction Stop | Out-Null
                    }
                } catch { L "  Provisioned package: $($_.Exception.Message)" }
                L '  Done'
            }
            L 'Finished removing selected apps.'
        }
        'install' {
            foreach ($it in $items) {
                $pkg, $disp = $it -split '=', 2
                L ">>> Installing $disp"
                $ok = $false
                foreach ($p in @(Get-AppxPackage -AllUsers -Name $pkg -ErrorAction SilentlyContinue)) {
                    $m = Join-Path $p.InstallLocation 'AppxManifest.xml'
                    if (Test-Path $m) {
                        try { Add-AppxPackage -DisableDevelopmentMode -Register $m -ErrorAction Stop; $ok = $true; L '  Registered from local manifest.' }
                        catch { L "  Manifest registration failed: $($_.Exception.Message)" }
                    }
                }
                if (-not $ok) {
                    L '  No local manifest - opening the Microsoft Store...'
                    try { Start-Process ('ms-windows-store://search/?query=' + [uri]::EscapeDataString($disp)) } catch { L "  FAILED: $($_.Exception.Message)" }
                }
            }
            L 'Finished installing selected apps.'
        }
    }
}

function Invoke-AppxAction([string]$mode) {
    if ($script:Busy) { Msg 'Another task is still running.' 'OK' 'Warning'; return }
    $sel = @($script:AppxChecks | Where-Object { $_.IsChecked } | ForEach-Object { $_.Tag })
    if ($sel.Count -eq 0) { Msg 'Select at least one app first.'; return }
    $verb = 'Install'; if ($mode -eq 'remove') { $verb = 'Remove' }
    if ((Msg "$verb $($sel.Count) selected app(s)?" 'YesNo' 'Question') -ne 'Yes') { return }
    $idStr = ($sel | ForEach-Object { '{0}={1}' -f $_.Pkg, $_.Name }) -join ';'
    $expLog.IsExpanded = $true
    Start-BG -Script $script:AppxJob -ArgList @($mode, $idStr) -Label "$verb selected apps..."
}

$btnAppRemoval.Add_Click({ Show-AppRemoval })
$btnAppBack.Add_Click({ $navTweaks.IsChecked = $true })
$btnAppDefault.Add_Click({ foreach ($cb in $script:AppxChecks) { $cb.IsChecked = [bool]$cb.Tag.Def } })
$btnAppAll.Add_Click({ foreach ($cb in $script:AppxChecks) { $cb.IsChecked = $true } })
$btnAppClear.Add_Click({ foreach ($cb in $script:AppxChecks) { $cb.IsChecked = $false } })
$btnAppGet.Add_Click({
    $done = {
        param($res)
        $names = @($res | ForEach-Object { [string]$_ })
        $n = 0
        foreach ($cb in $script:AppxChecks) {
            $cb.IsChecked = ($names -contains [string]$cb.Tag.Pkg)
            if ($cb.IsChecked) { $n++ }
        }
        Write-UILog "Get Installed: $n installed app(s) detected and selected."
    }
    Start-BG -Script $script:AppxJob -ArgList @('detect', '') -OnDone $done -Label 'Detecting installed apps...'
})
$btnAppInstall.Add_Click({ Invoke-AppxAction 'install' })
$btnAppRemove.Add_Click({ Invoke-AppxAction 'remove' })

# ============================================================================
#  UPDATES TAB
# ============================================================================
function Set-UpdateProfile([string]$p) {
    $wu = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'; $au = "$wu\AU"
    $tasks = @(Get-ScheduledTask -TaskPath '\Microsoft\Windows\WindowsUpdate\*' -ErrorAction SilentlyContinue) + @(Get-ScheduledTask -TaskPath '\Microsoft\Windows\UpdateOrchestrator\*' -ErrorAction SilentlyContinue)
    function SvcDefaults { Set-Service wuauserv -StartupType Manual -ErrorAction SilentlyContinue; Set-Service UsoSvc -StartupType Automatic -ErrorAction SilentlyContinue; Set-Service bits -StartupType Manual -ErrorAction SilentlyContinue }
    switch ($p) {
        'rec' {
            Remove-RegValue $au 'NoAutoUpdate'
            Set-Reg $wu 'DeferFeatureUpdates' 1; Set-Reg $wu 'DeferFeatureUpdatesPeriodInDays' 365
            Set-Reg $wu 'DeferQualityUpdates' 1; Set-Reg $wu 'DeferQualityUpdatesPeriodInDays' 4
            Set-Reg $wu 'ExcludeWUDriversInQualityUpdate' 1
            Set-Reg $au 'NoAutoRebootWithLoggedOnUsers' 1
            SvcDefaults; $tasks | ForEach-Object { Enable-ScheduledTask -InputObject $_ -ErrorAction SilentlyContinue | Out-Null }
        }
        'def' {
            Remove-Item $wu -Recurse -Force -ErrorAction SilentlyContinue
            SvcDefaults; $tasks | ForEach-Object { Enable-ScheduledTask -InputObject $_ -ErrorAction SilentlyContinue | Out-Null }
            Start-Service wuauserv -ErrorAction SilentlyContinue
        }
        'dis' {
            Set-Reg $au 'NoAutoUpdate' 1; Set-Reg $au 'AUOptions' 1
            foreach ($s in 'wuauserv', 'UsoSvc') { Stop-Service $s -Force -ErrorAction SilentlyContinue; Set-Service $s -StartupType Disabled -ErrorAction SilentlyContinue }
            $tasks | ForEach-Object { Disable-ScheduledTask -InputObject $_ -ErrorAction SilentlyContinue | Out-Null }
            Remove-Item "$env:SystemRoot\SoftwareDistribution\Download\*" -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}
$btnUpdRec.Add_Click({ Set-UpdateProfile 'rec'; Msg 'Recommended profile applied. Restart Windows to finish.' })
$btnUpdDef.Add_Click({ Set-UpdateProfile 'def'; Msg 'Windows default update behaviour restored. Restart Windows to finish.' })
$btnUpdDis.Add_Click({
    if ((Msg "Disabling updates means security patches will NOT be installed.`n`nContinue?" 'YesNo' 'Warning') -eq 'Yes') {
        Set-UpdateProfile 'dis'; Msg 'Updates disabled. Use "Restore Defaults" to turn them back on.'
    }
})

# ============================================================================
#  WINDOWS 10 / 11 ISO CREATOR TABS
# ============================================================================
$script:CreatorJob = {
    param($IsoPath, $OutPath, $Index, $O)
    $ErrorActionPreference = 'Stop'
    function L($m) { $Log.Enqueue(('[{0}] {1}' -f (Get-Date -Format 'HH:mm:ss'), $m)) }
    function Find-Oscdimg {
        $roots = @("${env:ProgramFiles(x86)}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools", "$env:ProgramFiles\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools")
        foreach ($r in $roots) { foreach ($a in 'amd64', 'x86', 'arm64') { $p = "$r\$a\Oscdimg\oscdimg.exe"; if (Test-Path $p) { return $p } } }
        $null
    }
    function RegAdd($key, $name, $val, $type = 'REG_DWORD') { & reg.exe add $key /v $name /t $type /d $val /f | Out-Null }

    $work = Join-Path $env:SystemDrive ('SysForge_Win{0}_{1}' -f $O.Ver, (Get-Date -Format 'yyyyMMddHHmmss'))
    $isoDir = Join-Path $work 'iso'; $mnt = Join-Path $work 'mount'; $newWim = Join-Path $work 'install_new.wim'
    $isoMounted = $false; $imgMounted = $false; $hives = @(); $ok = $false
    try {
        L "Starting Windows $($O.Ver) ISO creation."
        $free = (Get-PSDrive ($env:SystemDrive.TrimEnd(':'))).Free
        if ($free -lt 20GB) { throw 'At least 20 GB of free space is required on the system drive.' }

        $osc = Find-Oscdimg
        if (-not $osc) {
            L 'oscdimg.exe not found. Installing Windows ADK Deployment Tools through WinGet (one time, may take a while)...'
            if (Get-Command winget -ErrorAction SilentlyContinue) {
                & winget install -e --id Microsoft.WindowsADK --silent --accept-package-agreements --accept-source-agreements --override '/quiet /norestart /features OptionId.DeploymentTools' 2>&1 | ForEach-Object { $t = ([string]$_).Trim(); if ($t -and $t -notmatch '^[\-\\|/]+$') { L $t } }
                $osc = Find-Oscdimg
            }
        }
        if (-not $osc) { throw 'oscdimg.exe is required to build the ISO. Install the "Deployment Tools" feature of the Windows ADK and try again.' }

        New-Item -ItemType Directory $isoDir, $mnt -Force | Out-Null
        L 'Mounting source ISO...'
        Mount-DiskImage -ImagePath $IsoPath -StorageType ISO | Out-Null; $isoMounted = $true
        $drv = ((Get-DiskImage -ImagePath $IsoPath) | Get-Volume).DriveLetter
        L "Copying ISO contents from ${drv}: (several minutes)..."
        & robocopy.exe "${drv}:\" $isoDir /E /NFL /NDL /NJH /NJS /NP /R:1 /W:1 | Out-Null
        if ($LASTEXITCODE -ge 8) { throw 'Copying ISO contents failed.' }
        Dismount-DiskImage -ImagePath $IsoPath | Out-Null; $isoMounted = $false
        & attrib.exe -R "$isoDir\*" /S /D | Out-Null

        $src = Join-Path $isoDir 'sources\install.wim'
        if (-not (Test-Path $src)) { $src = Join-Path $isoDir 'sources\install.esd' }
        L "Exporting selected edition (index $Index) to a fresh install.wim..."
        Export-WindowsImage -SourceImagePath $src -SourceIndex $Index -DestinationImagePath $newWim -CompressionType Maximum | Out-Null

        L 'Mounting install image for offline servicing...'
        Mount-WindowsImage -ImagePath $newWim -Index 1 -Path $mnt | Out-Null; $imgMounted = $true

        if ($O.Apps) {
            L 'Removing preinstalled bloatware...'
            $pat = 'Microsoft.BingNews','Microsoft.BingWeather','Microsoft.GetHelp','Microsoft.Getstarted','Microsoft.Microsoft3DViewer','Microsoft.MicrosoftOfficeHub',
                   'Microsoft.MicrosoftSolitaireCollection','Microsoft.MixedReality.Portal','Microsoft.People','Microsoft.SkypeApp','Microsoft.WindowsFeedbackHub',
                   'Microsoft.YourPhone','Microsoft.ZuneMusic','Microsoft.ZuneVideo','Microsoft.Todos','Microsoft.PowerAutomateDesktop','Microsoft.WindowsMaps',
                   'Microsoft.Office.OneNote','Clipchamp.Clipchamp','Microsoft.549981C3F5F10','MicrosoftTeams','Microsoft.Windows.DevHome','Microsoft.BingSearch','Microsoft.OutlookForWindows'
            Get-AppxProvisionedPackage -Path $mnt | Where-Object { $pat -contains $_.DisplayName } | ForEach-Object {
                L "  - $($_.DisplayName)"; Remove-AppxProvisionedPackage -Path $mnt -PackageName $_.PackageName -ErrorAction SilentlyContinue | Out-Null
            }
        }

        if ($O.Tele -or $O.Oobe) {
            L 'Applying offline registry changes...'
            & reg.exe load HKLM\zSOFT "$mnt\Windows\System32\config\SOFTWARE" | Out-Null; if ($LASTEXITCODE -ne 0) { throw 'Could not load SOFTWARE hive.' }; $hives += 'HKLM\zSOFT'
            & reg.exe load HKLM\zUSER "$mnt\Users\Default\ntuser.dat" | Out-Null;           if ($LASTEXITCODE -ne 0) { throw 'Could not load default user hive.' }; $hives += 'HKLM\zUSER'
            if ($O.Tele) {
                RegAdd 'HKLM\zSOFT\Policies\Microsoft\Windows\DataCollection' 'AllowTelemetry' 0
                RegAdd 'HKLM\zSOFT\Policies\Microsoft\Windows\CloudContent' 'DisableWindowsConsumerFeatures' 1
                $cdm = 'HKLM\zUSER\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager'
                foreach ($n in 'SilentInstalledAppsEnabled','SystemPaneSuggestionsEnabled','ContentDeliveryAllowed','OemPreInstalledAppsEnabled','PreInstalledAppsEnabled','SubscribedContent-338388Enabled','SubscribedContent-338389Enabled','SubscribedContent-353698Enabled') { RegAdd $cdm $n 0 }
                RegAdd 'HKLM\zUSER\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo' 'Enabled' 0
            }
            if ($O.Oobe) {
                RegAdd 'HKLM\zSOFT\Policies\Microsoft\Windows\OOBE' 'DisablePrivacyExperience' 1
                if ($O.Ver -eq '11') { RegAdd 'HKLM\zSOFT\Microsoft\Windows\CurrentVersion\OOBE' 'BypassNRO' 1 }
            }
        }
        if ($O.OneDrive) {
            L 'Removing OneDrive setup...'
            foreach ($f in "$mnt\Windows\System32\OneDriveSetup.exe", "$mnt\Windows\SysWOW64\OneDriveSetup.exe") {
                if (Test-Path $f) { & takeown.exe /f $f | Out-Null; & icacls.exe $f /grant 'Administrators:F' | Out-Null; Remove-Item $f -Force -ErrorAction SilentlyContinue }
            }
            if ($hives -notcontains 'HKLM\zUSER') { & reg.exe load HKLM\zUSER "$mnt\Users\Default\ntuser.dat" | Out-Null; $hives += 'HKLM\zUSER' }
            & reg.exe delete 'HKLM\zUSER\Software\Microsoft\Windows\CurrentVersion\Run' /v OneDriveSetup /f 2>&1 | Out-Null
        }

        foreach ($h in $hives) { for ($i = 0; $i -lt 5; $i++) { [gc]::Collect(); Start-Sleep 1; & reg.exe unload $h 2>&1 | Out-Null; if ($LASTEXITCODE -eq 0) { break } } }
        $hives = @()

        L 'Saving changes and unmounting image (this can take a while)...'
        Dismount-WindowsImage -Path $mnt -Save | Out-Null; $imgMounted = $false

        Remove-Item (Join-Path $isoDir 'sources\install.esd') -Force -ErrorAction SilentlyContinue
        Remove-Item (Join-Path $isoDir 'sources\install.wim') -Force -ErrorAction SilentlyContinue
        Move-Item $newWim (Join-Path $isoDir 'sources\install.wim') -Force

        L 'Building bootable ISO with oscdimg...'
        $etfs = "$isoDir\boot\etfsboot.com"; $efi = "$isoDir\efi\microsoft\boot\efisys.bin"
        if ((Test-Path $etfs) -and (Test-Path $efi)) { $boot = "-bootdata:2#p0,e,b$etfs#pEF,e,b$efi" }
        elseif (Test-Path $efi) { $boot = "-bootdata:1#pEF,e,b$efi" }
        else { throw 'Boot files were not found in the copied ISO.' }
        $ErrorActionPreference = 'Continue'   # oscdimg prints progress on stderr; do not treat it as a terminating error
        & $osc -m -o -u2 -udfver102 "-lWIN$($O.Ver)" $boot $isoDir $OutPath 2>&1 | ForEach-Object { $t = ([string]$_).Trim(); if ($t -and $t -notmatch '^\d+%') { L $t } }
        $oscExit = $LASTEXITCODE
        $ErrorActionPreference = 'Stop'
        if ($oscExit -ne 0 -or -not (Test-Path $OutPath)) { throw "oscdimg failed to create the ISO (exit code $oscExit)." }
        $ok = $true
        L "SUCCESS: ISO saved to $OutPath"
    }
    catch { L "ERROR: $($_.Exception.Message)" }
    finally {
        L 'Cleaning up temporary files...'
        try {
            foreach ($h in $hives) { & reg.exe unload $h 2>&1 | Out-Null }
            if ($imgMounted) { Dismount-WindowsImage -Path $mnt -Discard | Out-Null }
            if ($isoMounted) { Dismount-DiskImage -ImagePath $IsoPath | Out-Null }
            if (Test-Path $work) { Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue }
        } catch { }
        if ($ok) { L 'All done.' } else { L 'Creation did not complete - see the messages above.' }
    }
}

$script:Cr = @{}
function New-CreatorView([string]$ver, [string]$url) {
    $root = [Windows.Markup.XamlReader]::Parse($creatorXaml.Replace('@@V@@', $ver))
    $c = @{ Root = $root; Ver = $ver; Url = $url }
    foreach ($n in 'btnBrowse', 'txtIso', 'step2', 'cmbEdition', 'chkApps', 'chkTele', 'chkOobe', 'chkOneDrive', 'btnCreate', 'btnOpenMs', 'txtStatus') { $c[$n] = $root.FindName($n) }
    [void]$contentHost.Children.Add($root)
    $script:Cr[$ver] = $c

    $c.btnOpenMs.Add_Click({ param($s, $e) Start-Process $s.Tag }.GetNewClosure())
    $c.btnOpenMs.Tag = $url
    $c.btnBrowse.Tag = $ver
    $c.btnBrowse.Add_Click({ param($s, $e) Select-Iso ([string]$s.Tag) })
    $c.btnCreate.Tag = $ver
    $c.btnCreate.Add_Click({ param($s, $e) Start-Creator ([string]$s.Tag) })
    $root
}

function Write-Creator([string]$ver, [string]$text) {
    $t = $script:Cr[$ver].txtStatus
    $t.AppendText(('[{0}] {1}{2}' -f (Get-Date -Format 'HH:mm:ss'), $text, [Environment]::NewLine)); $t.ScrollToEnd()
}

function Select-Iso([string]$ver) {
    $c = $script:Cr[$ver]
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = 'Windows ISO image (*.iso)|*.iso'; $dlg.Title = "Select official Windows $ver ISO"
    if (-not $dlg.ShowDialog()) { return }
    $path = $dlg.FileName
    $c.txtIso.Text = $path; $c.step2.Visibility = 'Collapsed'
    Write-Creator $ver "Selected: $path"
    Write-Creator $ver 'Reading editions from the ISO...'
    $window.Cursor = [Windows.Input.Cursors]::Wait; Pump
    $mountedHere = $false
    try {
        Mount-DiskImage -ImagePath $path -StorageType ISO | Out-Null; $mountedHere = $true
        $drv = ((Get-DiskImage -ImagePath $path) | Get-Volume).DriveLetter
        $wim = "${drv}:\sources\install.wim"; if (-not (Test-Path $wim)) { $wim = "${drv}:\sources\install.esd" }
        if (-not (Test-Path $wim)) { throw 'No install.wim / install.esd found. This is not a Windows installation ISO.' }
        $imgs = @(Get-WindowsImage -ImagePath $wim)
        $c.cmbEdition.Items.Clear(); $sel = 0; $k = 0
        foreach ($i in $imgs) {
            [void]$c.cmbEdition.Items.Add(('{0} - {1}' -f $i.ImageIndex, $i.ImageName))
            if ($i.ImageName -match 'Windows \d+ Pro$') { $sel = $k }
            $k++
        }
        $c.cmbEdition.SelectedIndex = $sel
        $c.step2.Visibility = 'Visible'
        Write-Creator $ver "Found $($imgs.Count) edition(s). Choose edition and options, then click Create."
    }
    catch { Write-Creator $ver "ERROR: $($_.Exception.Message)" }
    finally {
        if ($mountedHere) { Dismount-DiskImage -ImagePath $path -ErrorAction SilentlyContinue | Out-Null }
        $window.Cursor = $null
    }
}

function Start-Creator([string]$ver) {
    $c = $script:Cr[$ver]
    if ($script:Busy) { Msg 'Another task is still running.' 'OK' 'Warning'; return }
    $iso = $c.txtIso.Text
    if (-not (Test-Path $iso)) { Msg 'Please select a Windows ISO first.' 'OK' 'Warning'; return }
    if (-not $c.cmbEdition.SelectedItem) { Msg 'Please choose an edition.' 'OK' 'Warning'; return }
    $idx = [int](([string]$c.cmbEdition.SelectedItem -split ' - ')[0])
    $sfd = New-Object Microsoft.Win32.SaveFileDialog
    $sfd.Filter = 'ISO image (*.iso)|*.iso'; $sfd.FileName = "Win${ver}_SysForge.iso"; $sfd.Title = 'Save the new ISO as'
    if (-not $sfd.ShowDialog()) { return }
    if ((Msg "This will build a modified Windows $ver ISO.`nIt needs about 20 GB of free space on the system drive and can take 20-40 minutes.`n`nContinue?" 'YesNo' 'Question') -ne 'Yes') { return }
    $opts = @{
        Ver = $ver; Apps = [bool]$c.chkApps.IsChecked; Tele = [bool]$c.chkTele.IsChecked
        Oobe = [bool]$c.chkOobe.IsChecked; OneDrive = [bool]$c.chkOneDrive.IsChecked
    }
    Write-Creator $ver 'Starting background build...'
    Start-BG -Script $script:CreatorJob -ArgList @($iso, $sfd.FileName, $idx, $opts) -Label "Building Windows $ver ISO..." -Targets @($c.txtStatus)
}

$null = New-CreatorView '11' 'https://www.microsoft.com/software-download/windows11'
$null = New-CreatorView '10' 'https://www.microsoft.com/software-download/windows10ISO'

# ============================================================================
#  NAVIGATION / THEME / STARTUP
# ============================================================================
$script:Views = [ordered]@{
    navInstall   = $viewInstall
    navTweaks    = $viewTweaks
    navConfig    = $viewConfig
    navUpdates   = $viewUpdates
    navCreator11 = $script:Cr['11'].Root
    navCreator10 = $script:Cr['10'].Root
}
function Show-View([string]$nav) {
    $viewAppRemoval.Visibility = 'Collapsed'
    foreach ($k in $script:Views.Keys) { if ($k -eq $nav) { $script:Views[$k].Visibility = 'Visible' } else { $script:Views[$k].Visibility = 'Collapsed' } }
    if ($nav -like 'navCreator*') { $expLog.Visibility = 'Collapsed' } else { $expLog.Visibility = 'Visible' }
}
foreach ($n in $script:Views.Keys) { (Get-Variable $n -Scope Script -ValueOnly).Add_Checked({ param($s, $e) Show-View $s.Name }) }

# Settings (gear) button -> dropdown menu: Auto / Light Mode / Dark Mode
function Apply-ThemeMode([string]$mode) {
    $script:ThemeMode = $mode
    $target = $mode
    if ($mode -eq 'Auto') {
        $target = 'Light'
        if ((Get-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' 'AppsUseLightTheme' 1) -eq 0) { $target = 'Dark' }
    }
    Set-Theme $target
    foreach ($k in $script:ThemeItems.Keys) { $script:ThemeItems[$k].IsChecked = ($k -eq $mode) }
}
$script:ThemeMenu = New-Object Windows.Controls.ContextMenu
$script:ThemeItems = @{}
foreach ($opt in @(@('Auto','Auto'), @('Light','Light Mode'), @('Dark','Dark Mode'))) {
    $mi = New-Object Windows.Controls.MenuItem
    $mi.Header = $opt[1]
    $mi.Tag = $opt[0]
    $mi.IsCheckable = $true
    $mi.Add_Click({ param($s, $e) Apply-ThemeMode ([string]$s.Tag) })
    [void]$script:ThemeMenu.Items.Add($mi)
    $script:ThemeItems[$opt[0]] = $mi
}

# ---- Settings menu extras: Import / Export / About / Documentation ----
$script:DocsUrl = 'https://winutil.christitus.com/'   # change this to your own documentation page if you have one

function Export-SysForgeConfig {
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = 'SysForge config (*.json)|*.json'; $dlg.FileName = 'SysForge_Config.json'; $dlg.Title = 'Export SysForge selections'
    if (-not $dlg.ShowDialog()) { return }
    try {
        $mgr = 'winget'; if ($rbChoco.IsChecked) { $mgr = 'choco' }
        $cfg = [ordered]@{
            Tool           = 'SysForge'
            Version        = 1
            PackageManager = $mgr
            Apps           = @($script:Cards       | Where-Object { $_.IsChecked } | ForEach-Object { [string]$_.Tag.Name })
            Tweaks         = @($script:TweakChecks | Where-Object { $_.IsChecked } | ForEach-Object { [string]$_.Tag.Name })
            AppxPackages   = @($script:AppxChecks  | Where-Object { $_.IsChecked } | ForEach-Object { [string]$_.Tag.Pkg })
            Features       = @($script:FeatChecks  | Where-Object { $_.IsChecked } | ForEach-Object { [string]$_.Tag.N })
        }
        $cfg | ConvertTo-Json -Depth 5 | Set-Content -Path $dlg.FileName -Encoding UTF8
        Set-Status "Exported selections to $($dlg.FileName)"
        Write-UILog "Exported selections to $($dlg.FileName)"; Flush-Log
    } catch { Msg "Could not export: $($_.Exception.Message)" 'OK' 'Error' }
}

function Import-SysForgeConfig {
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = 'SysForge config (*.json)|*.json|All files (*.*)|*.*'; $dlg.Title = 'Import SysForge selections'
    if (-not $dlg.ShowDialog()) { return }
    try {
        $cfg = Get-Content -Path $dlg.FileName -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($cfg.Tool -ne 'SysForge') { throw 'This file is not a SysForge configuration.' }
        $apps  = @($cfg.Apps);         $twks  = @($cfg.Tweaks)
        $appx  = @($cfg.AppxPackages); $feats = @($cfg.Features)
        if ($cfg.PackageManager -eq 'choco') { $rbChoco.IsChecked = $true } else { $rbWinget.IsChecked = $true }
        foreach ($cb in $script:Cards)       { $cb.IsChecked = ($apps  -contains [string]$cb.Tag.Name) }
        foreach ($cb in $script:TweakChecks) { $cb.IsChecked = ($twks  -contains [string]$cb.Tag.Name) }
        foreach ($cb in $script:AppxChecks)  { $cb.IsChecked = ($appx  -contains [string]$cb.Tag.Pkg) }
        foreach ($cb in $script:FeatChecks)  { $cb.IsChecked = ($feats -contains [string]$cb.Tag.N) }
        Update-Count
        Set-Status "Imported selections from $($dlg.FileName)"
        Write-UILog "Imported selections from $($dlg.FileName)"; Flush-Log
    } catch { Msg "Could not import: $($_.Exception.Message)" 'OK' 'Error' }
}

function Show-About {
    Msg ("SysForge - Windows Utility`n`nA Windows system management tool inspired by Chris Titus Tech's WinUtil.`n`n" +
         "Install apps (WinGet / Chocolatey), apply tweaks, configure Windows features, manage updates and build debloated Windows 10/11 ISOs.`n`n" +
         "Create a restore point before applying tweaks.") 'OK' 'Information'
}

function Open-Documentation {
    try { Start-Process $script:DocsUrl } catch { Msg "Could not open the documentation: $($_.Exception.Message)" 'OK' 'Error' }
}

[void]$script:ThemeMenu.Items.Add((New-Object Windows.Controls.Separator))
$miImport = New-Object Windows.Controls.MenuItem; $miImport.Header = 'Import'
$miImport.Add_Click({ Import-SysForgeConfig })
[void]$script:ThemeMenu.Items.Add($miImport)
$miExport = New-Object Windows.Controls.MenuItem; $miExport.Header = 'Export'
$miExport.Add_Click({ Export-SysForgeConfig })
[void]$script:ThemeMenu.Items.Add($miExport)
[void]$script:ThemeMenu.Items.Add((New-Object Windows.Controls.Separator))
$miAbout = New-Object Windows.Controls.MenuItem; $miAbout.Header = 'About'
$miAbout.Add_Click({ Show-About })
[void]$script:ThemeMenu.Items.Add($miAbout)
$miDocs = New-Object Windows.Controls.MenuItem; $miDocs.Header = 'Documentation'
$miDocs.Add_Click({ Open-Documentation })
[void]$script:ThemeMenu.Items.Add($miDocs)

$btnTheme.Add_Click({
    $script:ThemeMenu.PlacementTarget = $btnTheme
    $script:ThemeMenu.Placement = [Windows.Controls.Primitives.PlacementMode]::Bottom
    $script:ThemeMenu.IsOpen = $true
})
Apply-ThemeMode 'Auto'

$window.Add_Closing({
    param($s, $e)
    if ($script:Busy) {
        if ((Msg "A task is still running. Closing now may leave it unfinished.`n`nClose anyway?" 'YesNo' 'Warning') -ne 'Yes') { $e.Cancel = $true }
    }
})
$window.Add_Closed({ $timer.Stop() })

Write-UILog 'SysForge ready. Create a restore point before applying tweaks.'
Flush-Log
[void]$window.ShowDialog()
