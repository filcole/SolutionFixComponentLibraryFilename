# TEMPORARY: There's a bug that causes default command libraries to export with inconsistent filenames.
# We will rename the files using the appId in the canvas app metadata.
#
# Find any the metadata file of any canvas apps of type component library
# For each canvas app CanvasApps/*meta.xml found
#   Extract the AppId from the metadata filename
#   Rename the BackgroundImage and DocumentUri files to use found AppId
#   Update the meta.xml to point to these moved files

# NOTES/WARNINGS:
# 1. The power platform solution is proprietary, this may stop working at any time.
# 2. Do not report issues with solution import/export to Microsoft if you're using this script

param (
    [Parameter(Mandatory = $true, HelpMessage = "Enter the path to the unpacked solution")]
    [Alias("p", "path")]
    [string]$solutionfolder
)

Function CheckSolutionFolder {

    if (!( Test-Path -Path $solutionfolder -PathType Container)) {
        Write-Error "Could not find folder $solutionfolder"
        exit
    }
    
    # Resolve any relative folder provided to script to full pathname
    $solnxml = Resolve-Path $solutionfolder

    $solnxml = Join-Path $solnxml "Other"
    $solnxml = Join-Path $solnxml "Solution.xml"

    if (!( Test-Path -Path $solnxml -PathType Leaf)) {
        Write-Error "Not valid solution folder. $solnxml does not exist"
        exit
    }
}

Function RenameComponentLibraryFile([string]$appXml, [string]$solutionfolder, [string]$tag, [string]$regex, [string]$appVersionHex)
{
    $re = "<$tag>" + $regex + "</$tag>"

    if (!($appXml -match $re)) {
        Write-Host "Warning: Could not find matching regex: $re"
        return $appXml
    }

    if ($matches[2] -eq $appVersionHex) {
        # No need to rename
        return $appXml
    }

    $fileName = $matches[1]  + $matches[2] + $matches[3]
    $newFileName = $matches[1] + $appVersionHex + $matches[3]

    Write-Host "Renaming $fileName to $newFileName"
    
    $origFullPath = Join-Path $solutionfolder $fileName
    $newFullPath = Join-Path $solutionfolder $newFileName
    
    Move-Item $origFullPath $newFullPath

    $appXml = $appXml -replace "<$tag>$fileName</$tag>", "<$tag>$newFileName</$tag>"

    # Return updated content
    return $appXml
}

Function FixComponentLibraryFilenames() {
    # TEMPORARY: There's a bug that causes default command libraries to export with inconsistent filenames.
    # We will save the filename with a hex version of the time in the AppVersion, so that it
    # only changes when the app version changes.

    # Find any the metadata file of any canvas apps of type component library
    # For each canvas app CanvasApps/*meta.xml found
    #   Extract the AppId from the metadata filename
    #   Rename the BackgroundImage and DocumentUri files to use found AppId
    #   Update the meta.xml to point to these moved files

    $canvasFolder = Join-Path $solutionfolder "CanvasApps"
    
    if (!(Test-Path $canvasFolder)) {
        Write-Debug "Canvas app folder does not exist"
        return
    }
    
    Write-Host "Scanning for Component Libraries with unstable filenames" 

    $reAppId = '^.*_([0-9a-z]{5})\.meta\.xml$'

    # CanvasAppType 1 = App Component Libraries, see
    # https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/reference/canvasapp
    # only these types of canvas app exhibit the problem
    $canvasApps = Get-ChildItem -Path "$canvasFolder" -Recurse -File -Filter *.meta.xml | 
    Select-String '<CanvasAppType>1</CanvasAppType>' -List | 
    Select-Object Path, Filename
    $canvasApps | ForEach-Object {
        $metadata = $_

        Write-Host "Checking app $($metadata.FileName)"

        $appXml = Get-Content -path $metadata.Path -Raw
        $origAppXml = $appXml

        if (!($metadata.FileName -match $reAppId)) {
            Write-Host "Warning: Could not find AppId in filename: $($metadata.FileName)"
            return
        }
    
        $appVersionHex = $matches[1]

        $appXml = RenameComponentLibraryFile `
            $appXml $solutionfolder `
            "BackgroundImageUri" "(/CanvasApps/\S+)([0-9a-z]{5})(_BackgroundImageUri)" $appVersionHex

        $appXml = RenameComponentLibraryFile `
            $appXml $solutionfolder `
            "DocumentUri" "(/CanvasApps/\S+)([0-9a-z]{5})(_DocumentUri\.msapp)" $appVersionHex

        if ($origAppXml -eq $appXml) {
            return
        }

        Write-Host "Saving updated metadata in $($metadata.FileName)"

        # Save the updated component library metadata
        $appXml | Set-Content -NoNewLine -Path $metadata.Path
    }
}

Function CheckUnmanagedSolutionExists() {
    # Check that the unmanaged solution exists in this environment
    $solnListText = pac solution list --json
    try {
        $solnListJson = ConvertFrom-Json $solnListText -ErrorAction Stop;
    }
    catch {
        Write-Host "pac solution list returned: $solnListText"
        exit
    }

    $solsfound = $solnListJson | Where-Object { $_.IsManaged -eq $false -and $_.SolutionUniqueName -eq $soln } | Measure-Object
    if ($solsfound.Count -ne 1) {
        Write-Host "Could not find unmanaged solution $soln. Correct environment?"
        exit
    }
}

## MAIN PROGRAM

CheckSolutionFolder

$solutionfolder = Resolve-Path $solutionfolder
FixComponentLibraryFilenames
