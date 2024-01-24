#Written by drdweb
#Searches for AVHDs and AVHDX checkpoints on all fixed drives

#from and to mail addresses
$mailsuffix = '@contoso.com'

$fromaddress = $env:COMPUTERNAME + ('-checkpoint_report') + $mailsuffix

$toaddress = 'backupse@contoso.com'

#Enumerate all logical disks and output drive name
$FixedDrives= (Get-WMIObject -Class Win32_LogicalDisk | Where-Object {$_.DriveType -eq "3"}).Name

#Search each drive recursively for AVHDX and AVHD files
$DiffFiles= foreach ($Drive in $FixedDrives){
    Get-ChildItem -Path $Drive\* -Include @("*.AVHDX","*.AVHD")  -Recurse -ErrorAction SilentlyContinue
}

#Formatting for properties and an expression to change "Length" to "Size" and adjust the measurement suffix appropriately 
$Properties = @(
    'Extension'
    'FullName'
    @{
        Label = 'Size'
        Expression = {
            if ($_.Length -ge 1GB)
            {
                '{0:F2} GB' -f ($_.Length / 1GB)
            }
            elseif ($_.Length -ge 1MB)
            {
                '{0:F2} MB' -f ($_.Length / 1MB)
            }
            elseif ($_.Length -ge 1KB)
            {
                '{0:F2} KB' -f ($_.Length / 1KB)
            }
            else
            {
                '{0} bytes' -f $_.Length
            }
        }
    }
    'CreationTime'
    'LastAccessTime'
    'LastWriteTime'
)

#Exit if no matching files are found
if ($null -eq $DiffFiles) {
    Exit
} else {
    #Apply formatting to results 
    $Table= $DiffFiles | Select-Object -Property $Properties

    #HTML style for email body
    $style = "<style>BODY{font-family: Arial; font-size: 10pt;}"
    $style = $style + "TABLE{border: 1px solid black; border-collapse: collapse;}"
    $style = $style + "TH{border: 1px solid black; background: #dddddd; padding: 5px; }"
    $style = $style + "TD{border: 1px solid black; padding: 5px; }"
    $style = $style + "</style>"

    #Convert the table to HTML using with formatting then output to a string
    $Body= $Table | ConvertTo-Html -Head $style | Out-String 

    #Set the subject to list the computername and domain
    $Subject= "Checkpoint search result for $env:COMPUTERNAME at $env:USERDNSDOMAIN"

    #Send results with the name of the machine and domain in the subject
    Send-MailMessage -From $fromaddress -Subject $Subject -To $backupdistro -SmtpServer mail.dynamic-networks.net -Body $Body -BodyAsHtml
}

