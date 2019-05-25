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


$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '510,381'
$Form.text                       = "Form"
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

$sync.source          = $source
$sync.destination     = $destination
$sync.sourcebrowse    = $sourcebrowse
$sync.destbrowse      = $destbrowse
$sync.fileprocessed   = $filesprocessed
$sync.initialsize     = $initialsize
$sync.currentsize     = $currentsize
$sync.percentsaved    = $percentsaved
$sync.handbrakeconfig = $handbrakeconfig
$sync.progresstext    = $progresstext
$sync.progressbar     = $ProgressBar
$sync.start           = $StartButton

$folderbrowse = New-Object System.Windows.Forms.FolderBrowserDialog
$folderbrowse.ShowNewFolderButton = $true
$folderbrowse.RootFolder = 'MyComputer'

$Form.controls.AddRange(@($Label1,$sync.source,$Label2,$sync.destination,$sync.sourcebrowse,$sync.destbrowse,$sync.fileprocessed,$label3,
    $sync.initialsize,$sync.currentsize,$sync.percentsaved,$sync.progresstext,$sync.progressbar,$Label7,$sync.start,$sync.handbrakeconfig
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


$sync.start.Add_Click({
    $script:handbrake = [PowerShell]::Create().AddScript({
        
    })
    $runspace = [RunspaceFactory]::CreateRunspace()
    $runspace.ApartmentState = "STA"
    $runspace.ThreadOptions = "ReuseThread"
    $runspace.Open()
    $runspace.SessionStateProxy.SetVariable("sync", $sync)
    $script:handbrake.Runspace = $runspace
    $script:handbrake.BeginInvoke()
})





<#   4675
$location = Read-Host -Prompt 'Input Path to Videos'
 
$finished = $location + "\" + "finished";

mkdir $finished;

# Path to HandbrakeCLI
$handbrakecli = "" 

# What video format to search for. Can't figure out how to do multiple.
$filter = "*.mkv" 

# HandbrakeCLI arguments. -i and -o are handled elsewhere
$arguments = "-e nvenc_h265 --encoder-preset slow -q 22 -E copy --comb-detect=fast --decomb=bob" 
 
$sizel = 0
$filelist = Get-ChildItem $location -filter $filter -exclude "_*" -recurse
$num = $filelist | Measure-Object
$filecount = $num.count
$sizea = Get-ChildItem $location -recurse | Measure-Object -Sum Length
$sizeb = "{0:N3}" -f ($sizea.sum/1gb)
$i = 0;
ForEach ($file in $filelist)
{
    $i++;

    # The two lines below automatically create the input and output file paths for HandbrakeCLI. It is not recommended to change this unless you know what you're doing.
    $oldfile = $file.DirectoryName + "\" + $file.BaseName + $file.Extension;

    # Change ".mp4" to $file.Extension to keep original filetype
    $newfile = $finished + "\" + $file.BaseName + ".mp4"; 
 
	# Next 8 lines handle progress and file size calculations
    $progress = ($i / $filecount) * 100
    $progress = [Math]::Round($progress,2)
    $sizec = Get-ChildItem $location -recurse | Measure-Object -Sum Length
    $sized = "{0:N3}" -f ($sizec.sum/1gb)
    $saved = (($sizeb - $sizee) / $sizeb) * 100
    $saved = [Math]::Round($saved,2)
	$sizef = get-childitem $oldfile | measure-object -sum length
	$sizeg = "{0:N3}" -f ($sizef.sum/1gb)
 
    Clear-Host
    Write-Host -------------------------------------------------------------------------------
    Write-Host Handbrake Batch Encoding
    Write-Host "Processing    - $oldfile"
    Write-Host "Progress      - $i of $filecount"
    Write-Host "Initial size  - $sizeb GB"
    Write-Host "Current size  - $sizee GB"
    Write-Host "Percent saved - $saved%"
    Write-Host -------------------------------------------------------------------------------
 
    Start-Process $handbrakecli -ArgumentList " $arguments -i `"$oldfile`" -o `"$newfile`"" -Wait -WindowStyle minimized
	$sizeh = get-childitem $newfile | measure-object -sum length
	$sizei = "{0:N3}" -f ($sizeh.sum/1gb)
	$sizel = $sizel + ($sizeg - $sizei)
	$sizee = $sizeb - $sizel
}
$sizec = Get-ChildItem $finished -recurse | Measure-Object -Sum Length
$sized = "{0:N3}" -f ($sizec.sum/1gb)
$saved = (($sizeb - $sized) / $sizeb) * 100
$saved = [Math]::Round($saved,2)
clear-host
Write-Host ----------------------------
Write-Host Handbrake Batch Encoding
Write-Host "Files processed - $filecount"
Write-Host "Initial size    - $sizeb GB"
Write-Host "Current Size    - $sized GB"
Write-Host "Percent saved   - $saved%"
Write-Host ----------------------------
Write-Host       
Write-Host Press any key to close
cmd /c pause | out-null
#>





[void]$Form.ShowDialog()