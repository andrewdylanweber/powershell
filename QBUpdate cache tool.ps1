#QBUpdate Tool
#drdweb 2023

#env setup
$basepath = "C:\SC_Temp\QBUpdateTool\"
$disabledir = 'disable-template'
$enabledir = 'enable-template'

#create basepath and template directories
New-Item -Path $basepath -Name $disabledir -ItemType Directory -Force
New-Item -Path $basepath -Name $enabledir -ItemType Directory -Force

#Get current ACLs and then set the desired access for Everyone
$disablesetup = Get-Acl -Path ($basepath + $disabledir)
$everyonedenyAcl = "Everyone", "FullControl", "ContainerInherit, ObjectInherit", "None", "Deny"
$EvDeny = New-Object System.Security.AccessControl.FileSystemAccessRule($everyonedenyAcl)

#add new ACLs to set on the next section
$disablesetup.AddAccessRule($EvDeny)
$disablesetup.SetAccessRuleProtection($True, $False)  

#Set the deny everyone permission on the template folder
Set-Acl -AclObject $disablesetup -Path ($basepath + $disabledir)

#array of QB update cache locations
$qbversions = @(
    #20
    "C:\ProgramData\Intuit\QuickBooks 2020\Components\QBUpdateCache\SPatch"
    #21
    "C:\ProgramData\Intuit\QuickBooks 2021\Components\QBUpdateCache\SPatch"
    #22
    "C:\ProgramData\Intuit\QuickBooks 2022\Components\DownloadQB32\spatch"
    #ent20
    "C:\ProgramData\Intuit\QuickBooks Enterprise Solutions 20.0\Components\QBUpdateCache\EPatch"
    #ent21
    "C:\ProgramData\Intuit\QuickBooks Enterprise Solutions 21.0\Components\QBUpdateCache\EPatch"
    #ent22
    "C:\ProgramData\Intuit\QuickBooks Enterprise Solutions 22.0\Components\DownloadQB32\EPatch"
)

#functions to loop through enabling\disabling folder access
function Disable-QBUpdates {
    $disable = Get-Acl -Path ($basepath + $disabledir)
    #disable loop
    foreach ($qb in $qbversions) {
        icacls.exe $qb /reset
        Set-Acl -Path $qb -AclObject $disable
    }
}


function Enable-QBUpdates {
    $enable = Get-Acl -Path ($basepath + $enabledir)
    #enable loop
    foreach ($qb in $qbversions) {
        icacls.exe $qb /reset
        Set-Acl -Path $qb -AclObject $enable
    }
}

Disable-QBUpdates

Enable-QBUpdates
