# Make sure that 7Zip is actually installed
$7zPath = "$env:ProgramFiles\7-Zip\7z.exe"
if (-not (Test-Path -Path $7zPath -PathType Leaf)) {
    throw "The 7zip application, '$7zPath' was not found. Please install 7Zip."
}

# Go to the MediData directory
cd "$env:SystemDrive\MediData\"

# Destination
$archivePath = "$env:SystemDrive\MediData-backup"
if ( -not (Test-Path -Path $archivePath -PathType Container) ) {
    Write-Output "Creating a $($archivePath) to store temporary backups."
    mkdir $archivePath
    Write-Output " ... Done."
}

$archive = "$($archivePath)\MediData-$(Get-Date -f yyyyMMdd-HHmmss).7z"

# Create the Archive
Set-Alias 7z $7zPath

Write-Output "Creating $($archive)";
7z a -mx=9 $archive -r -ssw
Write-Output " ... Done.";

# Move the newly created archive to the network path
$finalPath = "\\DocServer\Medisoft\AutoBackup\"
Write-Output "Looking for Remote Backup Location..."
if (Test-Path -Path "filesystem::$($finalPath)" -PathType Container) {
    Write-Output " ... Found!"
    Write-Output " ... Moving the Backup to the Remote Location"
    Move-Item -Path $archive -Destination $finalPath
    Write-Output " ... Done."

} else {
    Write-Output " ... Remote location is unavailable."
}

Write-Output "All Tasks Complete."