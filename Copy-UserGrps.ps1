#Written by drdweb 2-25-2020
function Copy-UserGrps {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$sourceUser,

        [Parameter(Mandatory = $true, Position = 1)]
        [string]$targetUser,

        [Parameter(Mandatory = $false, Position = 3)]
        [string]$excludedGroups = 'Domain Users'
    )

    begin {
        Import-Module ActiveDirectory -Force
        $srcCheck = Get-ADUser -Identity $sourceUser
        if ($null -eq $srcCheck) { Write-Host -ForegroundColor DarkRed "Could not find the source user account" }
        $grpSel = $srcCheck | Get-ADPrincipalGroupMembership | Where-Object { $_.Name -ne $excludedGroups } | Select-Object -Property Name, GroupScope, distinguishedName
        $trgtGrps = $grpSel.distinguishedName
    }

    process {
        foreach ($t in $trgtGrps) {
            Add-ADPrincipalGroupMembership -Identity $targetUser -MemberOf "$t" -ErrorAction SilentlyContinue
        }
    }

    end {
        $targetTest = (Get-ADPrincipalGroupMembership -Identity $targetUser).Count
        $sourceTest = (Get-ADPrincipalGroupMembership -Identity $sourceUser).Count

        if ($targetTest -eq $sourceTest) { 
            Write-Host "The group memberships of $targetUser and $SourceUser match!" -ForegroundColor Green
        }
        else {
            Write-Host "There is a mismatch in group memberships!"
        
            Write-Host $targetTest -ForegroundColor Yellow
            Write-Host $sourceTest -ForegroundColor Red
        }
        #Gets the current OU of the target user
        $dnsOU = (Get-ADUser -Identity $targetUser).DistinguishedName
        #Outputs the .NET listing of the source user's parent OU
        $parentOU = (Get-ADUser $sourceUser | Select-Object *, @{l = 'Parent'; e = { ([adsi]"LDAP://$($_.DistinguishedName)").Parent } }).Parent | Out-String
        #Trims the prefix of the parent OU to make it the correct syntax
        $targetOU = $parentOU.TrimStart("LDAP://")
        #Moves the target user to the OU of the source user
        Move-ADObject -Identity $dnsOU -TargetPath $targetOU
        Write-Host "$targetUser has been moved to the same OU as $sourceUser" -ForegroundColor DarkGreen

    }
}

<#Example Usage
Copy-UserGrps -sourceUser 'Administrator' -targetUser 'Webmin'
#>
