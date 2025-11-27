# PowerShell GUI to toggle CPU Boost Mode
# Efficient Enabled = 3, Aggressive = 2, Disabled = 0

# Check for administrator privileges
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
$isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    # Relaunch as administrator
    $scriptPath = $MyInvocation.MyCommand.Path
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`"" -Verb RunAs
    exit
}

Add-Type -AssemblyName System.Windows.Forms

# Function to get current boost mode
function Get-CurrentBoostMode {
    try {
        $scheme = (powercfg /getactivescheme) -replace '.*GUID: ([a-f0-9\-]+).*','$1'
        $subProcessor = "54533251-82be-4824-96c1-47b60b740d00"
        $boostMode = "be337238-0d82-4146-a960-4f3749d470c7"
        
        $acValue = powercfg /query $scheme $subProcessor $boostMode | Select-String "Current AC Power Setting Index:" | ForEach-Object { $_.Line -replace '.*: 0x([0-9a-fA-F]+).*','$1' }
        
        if ($acValue) {
            return [int]"0x$acValue"
        }
    }
    catch {
        # Return default if unable to read
        return 3
    }
    return 3
}

# Function to get processor performance settings
function Get-ProcessorSettings {
    try {
        $scheme = (powercfg /getactivescheme) -replace '.*GUID: ([a-f0-9\-]+).*','$1'
        $subProcessor = "54533251-82be-4824-96c1-47b60b740d00"
        $minProcState = "893dee8e-2bef-41e0-89c6-b55d0929964c"
        $maxProcState = "bc5038f7-23e0-4960-96da-33abaf5935ec"
        
        $minValue = powercfg /query $scheme $subProcessor $minProcState | Select-String "Current AC Power Setting Index:" | ForEach-Object { $_.Line -replace '.*: 0x([0-9a-fA-F]+).*','$1' }
        $maxValue = powercfg /query $scheme $subProcessor $maxProcState | Select-String "Current AC Power Setting Index:" | ForEach-Object { $_.Line -replace '.*: 0x([0-9a-fA-F]+).*','$1' }
        
        return @{
            Min = if ($minValue) { [int]"0x$minValue" } else { 5 }
            Max = if ($maxValue) { [int]"0x$maxValue" } else { 100 }
        }
    }
    catch {
        return @{ Min = 5; Max = 100 }
    }
}

# Function to unhide processor boost mode in Control Panel
function Enable-BoostModeVisibility {
    try {
        # Registry path for power settings
        $registryPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Power\PowerSettings\54533251-82be-4824-96c1-47b60b740d00\be337238-0d82-4146-a960-4f3749d470c7"
        
        # Check if the registry path exists
        if (Test-Path $registryPath) {
            # Set Attributes to 2 to make it visible in Control Panel
            # Attributes values: 0 or 2 = visible, 1 = hidden
            Set-ItemProperty -Path $registryPath -Name "Attributes" -Value 2 -Type DWord -ErrorAction Stop
            return $true
        }
        else {
            return $false
        }
    }
    catch {
        return $false
    }
}

# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = "CPU Performance Settings"
$form.Size = New-Object System.Drawing.Size(360,420)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false

# Create status label
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(20,15)
$lblStatus.Size = New-Object System.Drawing.Size(300,20)
$lblStatus.Text = "Current Mode: Loading..."
$lblStatus.Font = New-Object System.Drawing.Font("Arial",9,[System.Drawing.FontStyle]::Bold)

# Create groupbox for boost mode
$grpBoost = New-Object System.Windows.Forms.GroupBox
$grpBoost.Location = New-Object System.Drawing.Point(20,45)
$grpBoost.Size = New-Object System.Drawing.Size(310,115)
$grpBoost.Text = "CPU Boost Mode"

# Create radio buttons
$rbEfficient = New-Object System.Windows.Forms.RadioButton
$rbEfficient.Text = "Efficient Enabled (Balanced)"
$rbEfficient.Location = New-Object System.Drawing.Point(15,25)
$rbEfficient.Size = New-Object System.Drawing.Size(280,20)

$rbAggressive = New-Object System.Windows.Forms.RadioButton
$rbAggressive.Text = "Aggressive (Maximum Performance)"
$rbAggressive.Location = New-Object System.Drawing.Point(15,50)
$rbAggressive.Size = New-Object System.Drawing.Size(280,20)

$rbDisabled = New-Object System.Windows.Forms.RadioButton
$rbDisabled.Text = "Disabled (Power Saver)"
$rbDisabled.Location = New-Object System.Drawing.Point(15,75)
$rbDisabled.Size = New-Object System.Drawing.Size(280,20)

# Create groupbox for processor performance
$grpPerf = New-Object System.Windows.Forms.GroupBox
$grpPerf.Location = New-Object System.Drawing.Point(20,170)
$grpPerf.Size = New-Object System.Drawing.Size(310,130)
$grpPerf.Text = "Processor Performance (0-100%)"

# Min processor state label and value
$lblMinProc = New-Object System.Windows.Forms.Label
$lblMinProc.Location = New-Object System.Drawing.Point(15,25)
$lblMinProc.Size = New-Object System.Drawing.Size(150,20)
$lblMinProc.Text = "Minimum: 5%"

$trackMinProc = New-Object System.Windows.Forms.TrackBar
$trackMinProc.Location = New-Object System.Drawing.Point(15,45)
$trackMinProc.Size = New-Object System.Drawing.Size(280,45)
$trackMinProc.Minimum = 0
$trackMinProc.Maximum = 100
$trackMinProc.TickFrequency = 10
$trackMinProc.Value = 5

# Max processor state label and value
$lblMaxProc = New-Object System.Windows.Forms.Label
$lblMaxProc.Location = New-Object System.Drawing.Point(15,85)
$lblMaxProc.Size = New-Object System.Drawing.Size(150,20)
$lblMaxProc.Text = "Maximum: 100%"

$trackMaxProc = New-Object System.Windows.Forms.TrackBar
$trackMaxProc.Location = New-Object System.Drawing.Point(15,105)
$trackMaxProc.Size = New-Object System.Drawing.Size(280,45)
$trackMaxProc.Minimum = 0
$trackMaxProc.Maximum = 100
$trackMaxProc.TickFrequency = 10
$trackMaxProc.Value = 100

# Trackbar event handlers to update labels
$trackMinProc.Add_ValueChanged({
    $lblMinProc.Text = "Minimum: $($trackMinProc.Value)%"
})

$trackMaxProc.Add_ValueChanged({
    $lblMaxProc.Text = "Maximum: $($trackMaxProc.Value)%"
})

# Add controls to performance groupbox
$grpPerf.Controls.Add($lblMinProc)
$grpPerf.Controls.Add($trackMinProc)
$grpPerf.Controls.Add($lblMaxProc)
$grpPerf.Controls.Add($trackMaxProc)

# Create buttons
$btnApply = New-Object System.Windows.Forms.Button
$btnApply.Text = "Apply Settings"
$btnApply.Location = New-Object System.Drawing.Point(20,315)
$btnApply.Size = New-Object System.Drawing.Size(100,30)

$btnControlPanel = New-Object System.Windows.Forms.Button
$btnControlPanel.Text = "Show in CP"
$btnControlPanel.Location = New-Object System.Drawing.Point(130,315)
$btnControlPanel.Size = New-Object System.Drawing.Size(100,30)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Close"
$btnClose.Location = New-Object System.Drawing.Point(240,315)
$btnClose.Size = New-Object System.Drawing.Size(90,30)

# Event handler for Apply button
$btnApply.Add_Click({
    try {
        $scheme = (powercfg /getactivescheme) -replace '.*GUID: ([a-f0-9\-]+).*','$1'
        $subProcessor = "54533251-82be-4824-96c1-47b60b740d00"
        $boostMode    = "be337238-0d82-4146-a960-4f3749d470c7"
        $minProcState = "893dee8e-2bef-41e0-89c6-b55d0929964c"
        $maxProcState = "bc5038f7-23e0-4960-96da-33abaf5935ec"

        # Apply boost mode
        if ($rbEfficient.Checked) { 
            $value = 3 
            $modeName = "Efficient Enabled"
        }
        elseif ($rbAggressive.Checked) { 
            $value = 2 
            $modeName = "Aggressive"
        }
        elseif ($rbDisabled.Checked) { 
            $value = 0 
            $modeName = "Disabled"
        }
        else { 
            $value = 3
            $modeName = "Efficient Enabled"
        }

        powercfg /setacvalueindex $scheme $subProcessor $boostMode $value
        powercfg /setdcvalueindex $scheme $subProcessor $boostMode $value
        
        # Apply processor performance settings
        $minValue = $trackMinProc.Value
        $maxValue = $trackMaxProc.Value
        
        powercfg /setacvalueindex $scheme $subProcessor $minProcState $minValue
        powercfg /setdcvalueindex $scheme $subProcessor $minProcState $minValue
        powercfg /setacvalueindex $scheme $subProcessor $maxProcState $maxValue
        powercfg /setdcvalueindex $scheme $subProcessor $maxProcState $maxValue
        
        powercfg /S $scheme

        $lblStatus.Text = "Current Mode: $modeName"
        [System.Windows.Forms.MessageBox]::Show("Settings Applied:`n`nCPU Boost: $modeName`nMin Processor: $minValue%`nMax Processor: $maxValue%", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error applying settings: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Event handler for Control Panel button
$btnControlPanel.Add_Click({
    try {
        # Unhide the boost mode setting in registry
        $unhideResult = Enable-BoostModeVisibility
        
        if ($unhideResult) {
            # Open Power Options control panel
            Start-Process "control.exe" -ArgumentList "powercfg.cpl"
            
            # Show info message
            [System.Windows.Forms.MessageBox]::Show("Processor Boost Mode is now visible in Advanced Power Settings.`n`nTo access it:`n1. Click 'Change plan settings' for your active plan`n2. Click 'Change advanced power settings'`n3. Expand 'Processor power management'`n4. Expand 'Processor performance boost mode'", "Boost Mode Unhidden", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        else {
            # Still open Control Panel even if unhide failed
            Start-Process "control.exe" -ArgumentList "powercfg.cpl"
            [System.Windows.Forms.MessageBox]::Show("Control Panel opened, but boost mode setting may not be available on this system.`n`nThis could be because:`n- Your CPU doesn't support boost technology`n- The registry path doesn't exist`n- System limitations", "Note", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error opening Control Panel: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Event handler for Close button
$btnClose.Add_Click({
    $form.Close()
})

# Add controls to boost groupbox
$grpBoost.Controls.Add($rbEfficient)
$grpBoost.Controls.Add($rbAggressive)
$grpBoost.Controls.Add($rbDisabled)

# Add controls to form
$form.Controls.Add($lblStatus)
$form.Controls.Add($grpBoost)
$form.Controls.Add($grpPerf)
$form.Controls.Add($btnApply)
$form.Controls.Add($btnControlPanel)
$form.Controls.Add($btnClose)

# Load and set current boost mode
$currentMode = Get-CurrentBoostMode
switch ($currentMode) {
    3 { 
        $rbEfficient.Checked = $true
        $lblStatus.Text = "Current Mode: Efficient Enabled"
    }
    2 { 
        $rbAggressive.Checked = $true
        $lblStatus.Text = "Current Mode: Aggressive"
    }
    0 { 
        $rbDisabled.Checked = $true
        $lblStatus.Text = "Current Mode: Disabled"
    }
    default { 
        $rbEfficient.Checked = $true
        $lblStatus.Text = "Current Mode: Efficient Enabled"
    }
}

# Load and set processor performance settings
$procSettings = Get-ProcessorSettings
$trackMinProc.Value = $procSettings.Min
$trackMaxProc.Value = $procSettings.Max
$lblMinProc.Text = "Minimum: $($procSettings.Min)%"
$lblMaxProc.Text = "Maximum: $($procSettings.Max)%"

# Show form
$form.ShowDialog()