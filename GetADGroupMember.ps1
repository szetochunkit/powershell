function GetADGroupMember($group){
    $temp += (Get-ADGroupMember -Identity $group| where { $_.objectClass -eq "user" } |  select SamAccountName).SamAccountName
    #Get-ADGroup $group -Properties Member | Select-Object -Expand Member | Get-ADUser | where { $_.objectClass -eq "user" } | select samaccountname
    #$temp += (Get-ADGroup $group -Properties Member| Select-Object -Expand Member | Get-ADUser | select samaccountname).SamAccountName
    
    $result = @()
    foreach ( $i in $temp ){
        $result += [String]$i + "`n"
    }
    return [String]$result
}

foreach ( $group in $allgroups ){
    $groupname = $group.SamAccountName
    #$groupname
    $dn = (Get-ADGroup $groupname).DistinguishedName
    $output += [PSCustomObject]@{
        name = $groupname
        members = GetADGroupMember $groupname
        dn = $dn
    }
}
