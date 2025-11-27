Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Stalker Complete Mod Editor'
$form.Size = New-Object System.Drawing.Size(600, 400)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false

# Create labels and controls
$labelInfo = New-Object System.Windows.Forms.Label
$labelInfo.Location = New-Object System.Drawing.Point(10, 10)
$labelInfo.Size = New-Object System.Drawing.Size(560, 40)
$labelInfo.Text = 'This tool will modify weight limits in Stalker Complete mod files.' + [Environment]::NewLine + 'Select your mod folder to update system.ltx and actor.ltx'
$form.Controls.Add($labelInfo)

# Folder path textbox
$labelFolder = New-Object System.Windows.Forms.Label
$labelFolder.Location = New-Object System.Drawing.Point(10, 60)
$labelFolder.Size = New-Object System.Drawing.Size(100, 20)
$labelFolder.Text = 'Mod Folder:'
$form.Controls.Add($labelFolder)

$textboxFolder = New-Object System.Windows.Forms.TextBox
$textboxFolder.Location = New-Object System.Drawing.Point(10, 85)
$textboxFolder.Size = New-Object System.Drawing.Size(460, 20)
$textboxFolder.ReadOnly = $true
$form.Controls.Add($textboxFolder)

# Browse button
$buttonBrowse = New-Object System.Windows.Forms.Button
$buttonBrowse.Location = New-Object System.Drawing.Point(480, 83)
$buttonBrowse.Size = New-Object System.Drawing.Size(100, 25)
$buttonBrowse.Text = 'Browse...'
$form.Controls.Add($buttonBrowse)

# Status textbox (multiline)
$labelStatus = New-Object System.Windows.Forms.Label
$labelStatus.Location = New-Object System.Drawing.Point(10, 120)
$labelStatus.Size = New-Object System.Drawing.Size(100, 20)
$labelStatus.Text = 'Status:'
$form.Controls.Add($labelStatus)

$textboxStatus = New-Object System.Windows.Forms.TextBox
$textboxStatus.Location = New-Object System.Drawing.Point(10, 145)
$textboxStatus.Size = New-Object System.Drawing.Size(560, 150)
$textboxStatus.Multiline = $true
$textboxStatus.ScrollBars = 'Vertical'
$textboxStatus.ReadOnly = $true
$form.Controls.Add($textboxStatus)

# Apply changes button
$buttonApply = New-Object System.Windows.Forms.Button
$buttonApply.Location = New-Object System.Drawing.Point(200, 310)
$buttonApply.Size = New-Object System.Drawing.Size(120, 30)
$buttonApply.Text = 'Apply Changes'
$buttonApply.Enabled = $false
$form.Controls.Add($buttonApply)

# Close button
$buttonClose = New-Object System.Windows.Forms.Button
$buttonClose.Location = New-Object System.Drawing.Point(330, 310)
$buttonClose.Size = New-Object System.Drawing.Size(120, 30)
$buttonClose.Text = 'Close'
$form.Controls.Add($buttonClose)

# Browse button click event
$buttonBrowse.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = 'Select your Stalker Complete mod folder'
    $folderBrowser.ShowNewFolderButton = $false
    
    if ($folderBrowser.ShowDialog() -eq 'OK') {
        $textboxFolder.Text = $folderBrowser.SelectedPath
        $textboxStatus.Text = "Folder selected: $($folderBrowser.SelectedPath)" + [Environment]::NewLine
        
        # Determine the base path - if user selected gamedata folder, use it directly
        # Otherwise, look for gamedata subfolder
        $basePath = $folderBrowser.SelectedPath
        if ((Split-Path $basePath -Leaf) -eq 'gamedata') {
            # User selected the gamedata folder directly
            $systemLtx = Join-Path $basePath 'config\system.ltx'
            $actorLtx = Join-Path $basePath 'config\creatures\actor.ltx'
        } else {
            # User selected the parent folder containing gamedata
            $systemLtx = Join-Path $basePath 'gamedata\config\system.ltx'
            $actorLtx = Join-Path $basePath 'gamedata\config\creatures\actor.ltx'
        }
        
        $filesFound = $true
        
        if (Test-Path $systemLtx) {
            $textboxStatus.Text += "[OK] Found: system.ltx" + [Environment]::NewLine
        } else {
            $textboxStatus.Text += "[X] Missing: system.ltx" + [Environment]::NewLine
            $filesFound = $false
        }
        
        if (Test-Path $actorLtx) {
            $textboxStatus.Text += "[OK] Found: actor.ltx" + [Environment]::NewLine
        } else {
            $textboxStatus.Text += "[X] Missing: actor.ltx" + [Environment]::NewLine
            $filesFound = $false
        }
        
        if ($filesFound) {
            $textboxStatus.Text += [Environment]::NewLine + "Ready to apply changes!"
            $buttonApply.Enabled = $true
        } else {
            $textboxStatus.Text += [Environment]::NewLine + "Cannot proceed - missing required files."
            $buttonApply.Enabled = $false
        }
    }
})

# Apply changes button click event
$buttonApply.Add_Click({
    $folderPath = $textboxFolder.Text
    
    if ([string]::IsNullOrWhiteSpace($folderPath)) {
        [System.Windows.Forms.MessageBox]::Show('Please select a folder first.', 'Error', 'OK', 'Error')
        return
    }
    
    # Determine the base path - if user selected gamedata folder, use it directly
    # Otherwise, look for gamedata subfolder
    if ((Split-Path $folderPath -Leaf) -eq 'gamedata') {
        # User selected the gamedata folder directly
        $systemLtx = Join-Path $folderPath 'config\system.ltx'
        $actorLtx = Join-Path $folderPath 'config\creatures\actor.ltx'
    } else {
        # User selected the parent folder containing gamedata
        $systemLtx = Join-Path $folderPath 'gamedata\config\system.ltx'
        $actorLtx = Join-Path $folderPath 'gamedata\config\creatures\actor.ltx'
    }
    
    $textboxStatus.Text = "Starting modifications..." + [Environment]::NewLine
    $successCount = 0
    $errorCount = 0
    
    try {
        # Modify system.ltx
        if (Test-Path $systemLtx) {
            $textboxStatus.Text += "Processing system.ltx..." + [Environment]::NewLine
            
            # Create backup
            $backupPath = "$systemLtx.backup"
            Copy-Item $systemLtx $backupPath -Force
            $textboxStatus.Text += "  Backup created: system.ltx.backup" + [Environment]::NewLine
            
            # Read and modify content
            $content = Get-Content $systemLtx -Raw
            $originalContent = $content
            
            # Replace max_weight = 50 with max_weight = 1000
            $content = $content -replace '(max_weight\s*=\s*)50\b', '${1}1000'
            
            if ($content -ne $originalContent) {
                Set-Content $systemLtx -Value $content -NoNewline
                $textboxStatus.Text += "  [OK] Updated max_weight to 1000" + [Environment]::NewLine
                $successCount++
            } else {
                $textboxStatus.Text += "  [!] No changes needed (max_weight already set or not found)" + [Environment]::NewLine
            }
        }
        
        # Modify actor.ltx
        if (Test-Path $actorLtx) {
            $textboxStatus.Text += "Processing actor.ltx..." + [Environment]::NewLine
            
            # Create backup
            $backupPath = "$actorLtx.backup"
            Copy-Item $actorLtx $backupPath -Force
            $textboxStatus.Text += "  Backup created: actor.ltx.backup" + [Environment]::NewLine
            
            # Read and modify content
            $content = Get-Content $actorLtx -Raw
            $originalContent = $content
            
            # Replace max_item_mass = 50 with max_item_mass = 1000
            $content = $content -replace '(max_item_mass\s*=\s*)50\b', '${1}1000'
            
            # Replace max_walk_weight = 60 with max_walk_weight = 1000
            $content = $content -replace '(max_walk_weight\s*=\s*)60\b', '${1}1000'
            
            if ($content -ne $originalContent) {
                Set-Content $actorLtx -Value $content -NoNewline
                $textboxStatus.Text += "  [OK] Updated max_item_mass to 1000" + [Environment]::NewLine
                $textboxStatus.Text += "  [OK] Updated max_walk_weight to 1000" + [Environment]::NewLine
                $successCount++
            } else {
                $textboxStatus.Text += "  [!] No changes needed (values already set or not found)" + [Environment]::NewLine
            }
        }
        
        $textboxStatus.Text += [Environment]::NewLine + "════════════════════════════════════" + [Environment]::NewLine
        $textboxStatus.Text += "Modifications complete!" + [Environment]::NewLine
        $textboxStatus.Text += "Files modified: $successCount" + [Environment]::NewLine
        $textboxStatus.Text += "Backup files have been created." + [Environment]::NewLine
        
        [System.Windows.Forms.MessageBox]::Show('Changes applied successfully!', 'Success', 'OK', 'Information')
        
    } catch {
        $errorCount++
        $textboxStatus.Text += [Environment]::NewLine + "[X] ERROR: $($_.Exception.Message)" + [Environment]::NewLine
        [System.Windows.Forms.MessageBox]::Show("An error occurred: $($_.Exception.Message)", 'Error', 'OK', 'Error')
    }
})

# Close button click event
$buttonClose.Add_Click({
    $form.Close()
})

# Show the form
[void]$form.ShowDialog()
