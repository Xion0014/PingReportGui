# Ping Tool GUI Version Created by Caleb Flynn From Hypercare
# Refix Creation of LOG folder and dumping logs in there and NOT in DESKTOP

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# --- Helper: find the folder the script/exe lives in ---
function Get-BaseDirectory {
    if ($PSScriptRoot -and (Test-Path $PSScriptRoot)) { return $PSScriptRoot }
    if ($MyInvocation.MyCommand.Path) { return (Split-Path -Parent $MyInvocation.MyCommand.Path) }
    try {
        $procPath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
        if ($procPath -and (Test-Path $procPath)) { return (Split-Path -Parent $procPath) }
    } catch {}
    return (Get-Location).Path
}

# Create WPF window
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" 
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" 
        Title="Hypercare Store Ping Tool" Height="500" Width="800" WindowStartupLocation="CenterScreen" Background="#1E1E1E" Foreground="White">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <TextBlock Grid.Row="0" Text="CSV File Path:" Margin="0,0,0,5"/>
        <StackPanel Grid.Row="1" Orientation="Horizontal">
            <TextBox x:Name="csvPathBox" Width="600" Margin="0,0,5,0" Background="#2D2D30" Foreground="White"/>
            <Button x:Name="browseButton" Content="Browse..." Width="80"/>
        </StackPanel>
        
        <TextBlock Grid.Row="2" Text="Store Numbers (comma separated):" Margin="0,10,0,5"/>
        <TextBox x:Name="storeInput" Grid.Row="2" Margin="0,30,0,10" Height="25" Background="#2D2D30" Foreground="White"/>
        
        <ScrollViewer Grid.Row="3" VerticalScrollBarVisibility="Auto" Background="#1E1E1E" Margin="0,10,0,10">
            <RichTextBox x:Name="outputBox" IsReadOnly="True" Background="#1E1E1E" BorderThickness="0" FontFamily="Consolas" Foreground="White"/>
        </ScrollViewer>
        
        <Button x:Name="runButton" Grid.Row="4" Content="Run Ping Test" Height="35" Background="#007ACC" Foreground="White"/>
    </Grid>
</Window>
"@

# Read XAML
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$Window = [Windows.Markup.XamlReader]::Load($reader)

# Get controls
$csvPathBox = $Window.FindName("csvPathBox")
$browseButton = $Window.FindName("browseButton")
$storeInput = $Window.FindName("storeInput")
$outputBox = $Window.FindName("outputBox")
$runButton = $Window.FindName("runButton")

# Browse button
$browseButton.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "CSV Files (*.csv)|*.csv"
    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $csvPathBox.Text = $ofd.FileName
    }
})

# Function to append colored lines with custom font size and alignment
function Add-ColoredLine {
    param(
        $rtb,
        $storeNumber,
        $message,
        $color,
        [int]$fontSize = 12,
        [string]$alignment = "Left"
    )

    $brush = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb($color.R, $color.G, $color.B))
    $para = New-Object System.Windows.Documents.Paragraph
    $para.Margin = [Windows.Thickness]::new(0)
    $para.TextAlignment = $alignment
    $para.FontSize = $fontSize

    if ($storeNumber) {
        $storeRun = New-Object System.Windows.Documents.Run("$storeNumber - ")
        $storeRun.FontWeight = "Bold"
        $storeRun.Foreground = $brush
        $para.Inlines.Add($storeRun)
    }

    $msgRun = New-Object System.Windows.Documents.Run($message)
    $msgRun.FontWeight = "Bold"
    $msgRun.Foreground = $brush
    $para.Inlines.Add($msgRun)

    $rtb.Document.Blocks.Add($para)
    $rtb.ScrollToEnd()
}

# Thread-safe UI sync
$syncHash = [hashtable]::Synchronized(@{})
$syncHash.Window = $Window
$syncHash.LogAction = {
    param($storeNumber, $msg, $colorName, [int]$fontSize = 12, [string]$alignment = "Left")
    $syncHash.Window.Dispatcher.Invoke({
        switch ($colorName) {
            "Green"  { Add-ColoredLine $outputBox $storeNumber $msg ([System.Drawing.Color]::FromName("Lime")) $fontSize $alignment }
            "Red"    { Add-ColoredLine $outputBox $storeNumber $msg ([System.Drawing.Color]::FromName("Red")) $fontSize $alignment }
            "Yellow" { Add-ColoredLine $outputBox $storeNumber $msg ([System.Drawing.Color]::FromName("Yellow")) $fontSize $alignment }
            "Cyan"   { Add-ColoredLine $outputBox $storeNumber $msg ([System.Drawing.Color]::FromName("Cyan")) $fontSize $alignment }
            default  { Add-ColoredLine $outputBox $storeNumber $msg ([System.Drawing.Color]::FromName("White")) $fontSize $alignment }
        }
    })
}

# Run button
$runButton.Add_Click({
    $runButton.IsEnabled = $false
    $runButton.Content = "Running..."

    # --- FIX: Put logs in .\Logs next to the script/exe ---
    $baseDir = Get-BaseDirectory
    $logDir  = Join-Path $baseDir 'Logs'
    if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Force -Path $logDir | Out-Null }

    $globalCsvPath = Join-Path $logDir ("PingResults_{0}.csv" -f (Get-Date -Format 'yyyy-MM-dd_HHmmss'))
    if (Test-Path $globalCsvPath) { Remove-Item $globalCsvPath -Force }

    $storeList = $storeInput.Text.Split(',') | ForEach-Object { $_.Trim() }

    $ps = [powershell]::Create().AddScript({
        param($csvPath, $storeNumbers, $globalCsvPath, $syncHash)

        function Write-UI {
            param($storeNumber, $msg, $color, [int]$fontSize = 12, [string]$alignment = "Left")
            $syncHash.LogAction.Invoke($storeNumber, $msg, $color, $fontSize, $alignment)
        }

        try {
            if (-not (Test-Path $csvPath)) {
                Write-UI $null "ERROR: CSV file not found: $csvPath" "Red"
                return
            }

            # Ensure the LOG directory still exists (robustness for background thread)
            $outDir = Split-Path -Parent $globalCsvPath
            if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Force -Path $outDir | Out-Null }

            $dataRaw = Import-Csv -Path $csvPath
            $data = foreach ($row in $dataRaw) {
                $newRow = @{ }
                foreach ($col in $row.PSObject.Properties) {
                    $cleanName = $col.Name -replace ' ', ''
                    $newRow[$cleanName] = $col.Value
                }
                [PSCustomObject]$newRow
            }

            # CSV Header - updated order
            "IP,Store,Device,Status,ResponseTime (ms)" | Out-File -FilePath $globalCsvPath -Encoding UTF8

            foreach ($storeNumber in $storeNumbers) {
                # Large centered cyan header
                Write-UI $null "`n=== Running Store $storeNumber ===" "Cyan" 18 "Center"
                # Cyan horizontal line
                Write-UI $null "----------------------------------------" "Cyan" 12 "Center"

                $storeRow = $data | Where-Object { $_.Store -eq $storeNumber }
                if (-not $storeRow) {
                    Write-UI $storeNumber "ERROR: Store not found in CSV." "Red"
                    continue
                }

                $devices = $storeRow.PSObject.Properties | Where-Object { $_.Name -ne "Store" }
                $devicesSorted = $devices | Sort-Object {
                    if ($_.Value -match "\d+\.\d+\.\d+\.(\d+)$") { [int]$matches[1] } else { 999 }
                }

                foreach ($device in $devicesSorted) {
                    $deviceName = $device.Name
                    $ip = $device.Value.Trim()
                    if (-not $ip) {
                        Write-UI $storeNumber "${deviceName}: Skipped (no IP)" "Yellow"
                        ",$storeNumber,${deviceName},Skipped (no IP)," | Out-File -FilePath $globalCsvPath -Append
                        continue
                    }

                    try {
                        $ping = Test-Connection -ComputerName $ip -Count 1 -ErrorAction Stop
                        $time = $ping.ResponseTime
                        Write-UI $storeNumber "${deviceName} (${ip}): Success (${time} ms)" "Green"
                        "$ip,$storeNumber,${deviceName},Success,$time" | Out-File -FilePath $globalCsvPath -Append
                    } catch {
                        Write-UI $storeNumber "${deviceName} (${ip}): Failed" "Red"
                        "$ip,$storeNumber,${deviceName},Failed," | Out-File -FilePath $globalCsvPath -Append
                    }
                }
            }

            Write-UI $null "Results saved to $globalCsvPath" "Cyan"

        } catch {
            Write-UI $null "ERROR: $($_.Exception.Message)" "Red"
        }

        # Re-enable button and reset text
        $syncHash.Window.Dispatcher.Invoke({
            $btn = $syncHash.Window.FindName("runButton")
            $btn.IsEnabled = $true
            $btn.Content = "Run Ping Test"
        })

    }).AddArgument($csvPathBox.Text).AddArgument($storeList).AddArgument($globalCsvPath).AddArgument($syncHash)

    $ps.Runspace = [runspacefactory]::CreateRunspace()
    $ps.Runspace.ApartmentState = "STA"
    $ps.Runspace.ThreadOptions = "ReuseThread"
    $ps.Runspace.Open()
    $ps.BeginInvoke()
})

# Show window
$Window.ShowDialog() | Out-Null