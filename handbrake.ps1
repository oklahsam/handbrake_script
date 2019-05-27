Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

# Hide PowerShell Console
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
$consolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($consolePtr, 0)

# for talking across runspaces.
$sync = [Hashtable]::Synchronized(@{})
$config = [Hashtable]::Synchronized(@{})

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '510,381'
$Form.text                       = "Video Resize"
$Form.TopMost                    = $false
$form.maximizebox                = $false
$form.formborderstyle            = 'Fixed3d'

$source                          = New-Object system.Windows.Forms.TextBox
$source.multiline                = $false
$source.width                    = 414
$source.height                   = 30
$source.location                 = New-Object System.Drawing.Point(12,37)
$source.Font                     = 'Microsoft Sans Serif,10'

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "Source:"
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(12,17)
$Label1.Font                     = 'Microsoft Sans Serif,10'

$destination                     = New-Object system.Windows.Forms.TextBox
$destination.multiline           = $false
$destination.width               = 414
$destination.height              = 20
$destination.location            = New-Object System.Drawing.Point(12,100)
$destination.Font                = 'Microsoft Sans Serif,10'

$Label2                          = New-Object system.Windows.Forms.Label
$Label2.text                     = "Destination:"
$Label2.AutoSize                 = $true
$Label2.width                    = 25
$Label2.height                   = 10
$Label2.location                 = New-Object System.Drawing.Point(12,79)
$Label2.Font                     = 'Microsoft Sans Serif,10'

$sourcebrowse                    = New-Object system.Windows.Forms.Button
$sourcebrowse.text               = "Browse"
$sourcebrowse.width              = 60
$sourcebrowse.height             = 19
$sourcebrowse.location           = New-Object System.Drawing.Point(435,38)
$sourcebrowse.Font               = 'Microsoft Sans Serif,10'

$destbrowse                      = New-Object system.Windows.Forms.Button
$destbrowse.text                 = "Browse"
$destbrowse.width                = 60
$destbrowse.height               = 19
$destbrowse.location             = New-Object System.Drawing.Point(435,101)
$destbrowse.Font                 = 'Microsoft Sans Serif,10'

$filesprocessed                  = New-Object system.Windows.Forms.Label
$filesprocessed.text             = "Files Processed:"
$filesprocessed.AutoSize         = $true
$filesprocessed.width            = 25
$filesprocessed.height           = 10
$filesprocessed.location         = New-Object System.Drawing.Point(12,132)
$filesprocessed.Font             = 'Microsoft Sans Serif,10'

$initialsize                     = New-Object system.Windows.Forms.Label
$initialsize.text                = "Initial Size:"
$initialsize.AutoSize            = $true
$initialsize.width               = 25
$initialsize.height              = 10
$initialsize.location            = New-Object System.Drawing.Point(44,154)
$initialsize.Font                = 'Microsoft Sans Serif,10'

$currentsize                     = New-Object system.Windows.Forms.Label
$currentsize.text                = "Current Size:"
$currentsize.AutoSize            = $true
$currentsize.width               = 25
$currentsize.height              = 10
$currentsize.location            = New-Object System.Drawing.Point(33,176)
$currentsize.Font                = 'Microsoft Sans Serif,10'

$percentsaved                    = New-Object system.Windows.Forms.Label
$percentsaved.text               = "Percent Saved:"
$percentsaved.AutoSize           = $true
$percentsaved.width              = 25
$percentsaved.height             = 10
$percentsaved.location           = New-Object System.Drawing.Point(20,198)
$percentsaved.Font               = 'Microsoft Sans Serif,10'

$label3                          = New-Object system.Windows.Forms.Label
$label3.text                     = "Handbrake Config:"
$label3.AutoSize                 = $true
$label3.width                    = 25
$label3.height                   = 10
$label3.location                 = New-Object System.Drawing.Point(12,242)
$label3.Font                     = 'Microsoft Sans Serif,8'

$handbrakeconfig                 = New-Object system.Windows.Forms.TextBox
$handbrakeconfig.multiline       = $false
$handbrakeconfig.width           = 486
$handbrakeconfig.height          = 20
$handbrakeconfig.location        = New-Object System.Drawing.Point(12,260)
$handbrakeconfig.Font            = 'Microsoft Sans Serif,10'

$progresstext                    = New-Object system.Windows.Forms.TextBox
$progresstext.multiline          = $false
$progresstext.width              = 486
$progresstext.height             = 10
$progresstext.location           = New-Object System.Drawing.Point(12,287)
$progresstext.Font               = 'Microsoft Sans Serif,8'
$progresstext.ReadOnly           = $true

$ProgressBar                     = New-Object system.Windows.Forms.ProgressBar
$ProgressBar.width               = 486
$ProgressBar.height              = 24
$ProgressBar.location            = New-Object System.Drawing.Point(12,312)

$currentfile                     = New-Object system.Windows.Forms.Label
$currentfile.text                = "Current File:"
$currentfile.AutoSize            = $true
$currentfile.width               = 25
$currentfile.height              = 10
$currentfile.location            = New-Object System.Drawing.Point(12,271)
$currentfile.Font                = 'Microsoft Sans Serif,10'

$StartButton                     = New-Object system.Windows.Forms.Button
$StartButton.text                = "Start"
$StartButton.width               = 486
$StartButton.height              = 30
$StartButton.location            = New-Object System.Drawing.Point(12,340)
$StartButton.Font                = 'Microsoft Sans Serif,10'

$hbdefault                      = New-Object system.Windows.Forms.Button
$hbdefault.text                 = "Default"
$hbdefault.width                = 45
$hbdefault.height               = 17
$hbdefault.location             = New-Object System.Drawing.Point(453,240)
$hbdefault.Font                 = 'Microsoft Sans Serif,7'

$sync.source          = $source
$sync.destination     = $destination
$sync.sourcebrowse    = $sourcebrowse
$sync.destbrowse      = $destbrowse
$sync.filesprocessed  = $filesprocessed
$sync.initialsize     = $initialsize
$sync.currentsize     = $currentsize
$sync.percentsaved    = $percentsaved
$sync.handbrakeconfig = $handbrakeconfig
$sync.progresstext    = $progresstext
$sync.progressbar     = $ProgressBar
$sync.start           = $StartButton
$sync.hbdefault       = $hbdefault

$folderbrowse = New-Object System.Windows.Forms.FolderBrowserDialog
$folderbrowse.ShowNewFolderButton = $true
$folderbrowse.RootFolder = 'MyComputer'

$Form.controls.AddRange(@($Label1,$sync.source,$Label2,$sync.destination,$sync.sourcebrowse,$sync.destbrowse,$sync.filesprocessed,$label3,
    $sync.initialsize,$sync.currentsize,$sync.percentsaved,$sync.progresstext,$sync.progressbar,$Label7,$sync.start,$sync.handbrakeconfig,
    $sync.hbdefault
))

$sync.sourcebrowse.Add_Click({
    $folderbrowse.SelectedPath = ""
    $null = $Folderbrowse.ShowDialog()
    $sync.source.text = $Folderbrowse.SelectedPath
})

$sync.destbrowse.Add_Click({
    $folderbrowse.SelectedPath = ""
    $null = $Folderbrowse.ShowDialog()
    $sync.destination.text = $Folderbrowse.SelectedPath
})

$sync.hbdefault.add_click({
    $handbrakeconfig.text = "-e nvenc_h264  -q 22 -E copy --comb-detect=fast --decomb=bob"
})

$sync.start.Add_Click({
    $script:handbrake = [PowerShell]::Create().AddScript({
        if ( -not [String]::IsNullOrWhiteSpace($sync.destination.text) ) {
            Invoke-Expression $sync.handbrake
        } else {
            $sync.progresstext.text = "No destination selected"
            $script:handbrake.runspace.dispose()
            $script:handbrake.dispose()
        }
    })
    $runspace = [RunspaceFactory]::CreateRunspace()
    $runspace.ApartmentState = "STA"
    $runspace.ThreadOptions = "ReuseThread"
    $runspace.Open()
    $runspace.SessionStateProxy.SetVariable("sync", $sync)
    $script:handbrake.Runspace = $runspace
    $script:handbrake.BeginInvoke()
})

function HANDBRAKE {
    Invoke-Expression $sync.disablecontrols
    $count = 0
    $sync.progressbar.value = 0
    $files = get-childitem $sync.source.text -recurse -file -include *.mp4,*.mkv,*.avi,*.mov | Select-Object extension,fullName,basename, @{Name="Bytes";Expression={ "{0:N0}" -f ($_.Length / 1MB) }}
    $sourcetext = $sync.source.text
    if ($files.fullname.count -eq 0) {
        $files = get-childitem ($sync.source.text + "*\*") -recurse -file -include *.mp4,*.mkv,*.avi,*.mov | Select-Object directory,extension,fullName,basename, @{Name="Bytes";Expression={ "{0:N0}" -f ($_.Length / 1MB) }}
        $sourcetext = $sync.source.text.substring(0, $sync.source.text.lastindexof('\'))
    }
    if ($files.fullname.count -gt 0) {
        $config = $sync.handbrakeconfig.text
        $size = ($files | measure-object -property Bytes -sum).sum
        if (($size) -ge "10000") { 
            $start_size = [math]::round(($size / 1KB),3)
            $sync.initialsize.text = "Initial Size: " + $start_size + " GBs"
            $sync.currentsize.text = "Current Size: " + $start_size + " GBs"
        } else {
            $start_size = [math]::round(($size),2)
            $sync.initialsize.text = "Initial Size: " + $start_size + " MBs"
            $sync.currentsize.text = "Current Size: " + $start_size + " MBs"
        }
        $new_size = $start_size
        $sync.filesprocessed.text = "Files Processed: 0 out of " + $files.fullname.count
        $sync.percentsaved.text = "Percent Saved: 0%"
        foreach ($file in $files) {
            $count++
            $in = $file.fullname
            $sync.progresstext.text = $file.fullname
            $dest = $file.fullname -replace [regex]::Escape($sourcetext), $sync.destination.text -replace $file.Extension, ".mp4"
            if ((test-path $dest) -eq $false) {
                new-item $dest -force
                Start-Process "C:\handbrake\HandBrakeCLI.exe" -ArgumentList " $config -i `"$in`" -o `"$dest`"" -Wait -WindowStyle minimized
            }
            if (($size) -ge "10000") { 
                $new_file =  get-childitem $dest -recurse -file | Select-Object @{Name="Bytes";Expression={ "{0:0.###}" -f ($_.Length / 1GB) }}
                $new_file = $new_file.bytes
                $old_file = $file.bytes / 1KB
                $new_size = [math]::round(($new_size - ($old_file - $new_file)),3)
                $sync.currentsize.text = "Current Size: " + $new_size + " GBs"
            } else {
                $new_file =  get-childitem $dest -recurse -file | Select-Object @{Name="Bytes";Expression={ "{0:0.##}" -f ($_.Length / 1MB) }}
                $new_file = $new_file.bytes
                $old_file = $file.bytes
                $new_size = [math]::round(($new_size - ($old_file - $new_file)),2)
                $sync.currentsize.text = "Current Size: " + $new_size + " MBs"
            }
            $percent = [math]::round(((($start_size - $new_size) / $start_size) * 100),2)
            $sync.percentsaved.text = "Percent Saved: " + $percent + "%"
            $sync.filesprocessed.text = "Files Processed: " + $count + " out of " + $files.fullname.count
            $sync.progressbar.value = $sync.progressbar.value + (($old_file / $start_size) * 100)
        }
        $sync.progresstext.text = "Done"
    } else { $sync.progresstext.text = "No video files in source" }
    Invoke-Expression $sync.enablecontrols
    $script:handbrake.runspace.dispose()
    $script:handbrake.dispose()
}
$sync.handbrake = get-content Function:\HANDBRAKE

function ENABLECONTROLS {
    $sync.source.enabled          = $true
    $sync.destination.enabled     = $true
    $sync.sourcebrowse.enabled    = $true
    $sync.destbrowse.enabled      = $true
    $sync.start.enabled           = $true
    $sync.handbrakeconfig.enabled = $true
    $sync.hbdefault               = $true
}
$sync.enablecontrols = get-content Function:\ENABLECONTROLS

function DISABLECONTROLS {
    $sync.source.enabled          = $false
    $sync.destination.enabled     = $false
    $sync.sourcebrowse.enabled    = $false
    $sync.destbrowse.enabled      = $false
    $sync.start.enabled           = $false
    $sync.handbrakeconfig.enabled = $false
    $sync.hbdefault               = $false
}
$sync.disablecontrols = get-content Function:\DISABLECONTROLS

$config = import-clixml ($ENV:userprofile + "\handbrakeconfig.xml")
$sync.handbrakeconfig.text = $config.handbrake

[void]$Form.ShowDialog()

$config.handbrake = $sync.handbrakeconfig.Text
$config | export-clixml ($ENV:userprofile + "\handbrakeconfig.xml")

$script:handbrake.runspace.dispose()
$script:handbrake.dispose()