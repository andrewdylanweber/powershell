#Directory file list tool
Function Get-FolderName {
    [CmdletBinding()]
    [OutputType([string])]
    Param
    (
        # InitialDirectory help description
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            HelpMessage = "Initial Directory for browsing",
            Position = 0)]
        [String]$SelectedPath,

        # Description help description
        [Parameter(
            Mandatory = $false,
            ValueFromPipelineByPropertyName = $true,
            HelpMessage = "Message Box Title")]
        [String]$Description = "Select a Folder",

        # ShowNewFolderButton help description
        [Parameter(
            Mandatory = $false,
            HelpMessage = "Show New Folder Button when used")]
        [Switch]$ShowNewFolderButton
    )

    # Load Assembly
    Add-Type -AssemblyName System.Windows.Forms

    # Open Class
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog

    # Define Title
    $FolderBrowser.Description = $Description

    # Define Initial Directory
    if (-Not [String]::IsNullOrWhiteSpace($SelectedPath)) {
        $FolderBrowser.SelectedPath = $SelectedPath
    }

    if ($folderBrowser.ShowDialog() -eq "OK") {
        $Folder += $FolderBrowser.SelectedPath
    }
    return $Folder
}

function Save-FileName([string] $initialDirectory ) {

    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "Text Files (*.txt)| *.txt"
    $OpenFileDialog.ShowDialog() |  Out-Null

    return $OpenFileDialog.filename
}

$rootfolder = Get-FolderName -SelectedPath "My Computer"

$depth = '4'

$outputfile = Save-FileName


(Get-ChildItem -Path $rootfolder -Depth $depth | 
Select-Object -Property FullName).FullName |
Out-File -FilePath $outputfile
