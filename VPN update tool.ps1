#VPN updater
#Written by drdweb Oct 2022

#Variable for the VPN target
$L2TP_Public = 'vpn.contoso.com'
$VPNName = 'contoso-vpn'
$psk = 'supersecurestring1234!@#$'
#attempts to add a new VPN with the name $VPNName
try {
    Add-VpnConnection -AllUserConnection -Name $VPNName -ServerAddress $L2TP_Public -TunnelType L2tp -EncryptionLevel Required -L2tpPsk $psk -AuthenticationMethod MSCHAPv2 -RememberCredential $false -IdleDisconnectSeconds 900 -Force
}
#If it fails to add the vpn it attempts to update an existing config
catch {
    Set-VpnConnection -AllUserConnection -Name $VPNName -ServerAddress $L2TP_Public -TunnelType L2tp -EncryptionLevel Required -L2tpPsk $psk -AuthenticationMethod MSCHAPv2 -RememberCredential $false -IdleDisconnectSeconds 900 -Force
}






