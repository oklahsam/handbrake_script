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
    $newfile = $finished + "\" + $file.BaseName + ".mp4"; # Change ".mp4" to $file.Extension to keep original filetype
 
	# Next 6 lines handle progress and file size calculations
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