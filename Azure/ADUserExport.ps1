# Exports users in a given OU with fields ready to paste into the Master Calc Sheet

$SearchBase = "OU=Staff,OU=Accounts,OU=WFI,OU=South Farnham Education Trust,DC=ad,DC=sfet,DC=org,DC=uk"
$ExportPath = "C:\EDUIT\WFI_Staff_AD_User_Export.csv"

Get-ADUser -Filter * -SearchBase $SearchBase -Properties displayName, sAMAccountName, UserPrincipalName, GivenName, Surname, distinguishedName, enabled, homeDirectory, Profilepath, Pager, Description, mail, ProxyAddresses, ObjectGUID |
    Select-Object displayName, sAMAccountName, UserPrincipalName, GivenName, Surname, distinguishedName, enabled, homeDirectory, Profilepath, Pager, Description, mail,
        @{L='ProxyAddress_1'; E={$_.proxyaddresses[0]}}, 
        @{L='ProxyAddress_2'; E={$_.ProxyAddresses[1]}}, 
        @{L='ProxyAddress_3'; E={$_.ProxyAddresses[2]}}, 
        @{L='ProxyAddress_4'; E={$_.ProxyAddresses[3]}}, 
        @{L='ProxyAddress_5'; E={$_.ProxyAddresses[4]}},
        ObjectGUID |
    Sort-Object samAccountName | 
    Export-Csv $ExportPath -NoTypeInformation
