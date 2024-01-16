$all_dc = @()
$all_dc += Get-ADComputer -SearchBase "OU=Domain Controllers,DC=lexinfintech,DC=com" -Filter * | sort 

function GetOsTime( $servername ){
    # 使用 Get-WmiObject 获取系统时间
    try{
        $os = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $servername
        if( $? -eq $false ){
            return "error1"       
        }
    }catch{
        return "error2"
    }
    $originalString = $os.LocalDateTime

    # 格式化系统时间为年月日时分秒格式
    #$formattedDateTime = Get-Date -Date $localDateTime -Format "yyyy-MM-dd HH:mm:ss"
    
    # 从原始字符串中提取各个日期时间部分
    $year = $originalString.Substring(0, 4)
    $month = $originalString.Substring(4, 2)
    $day = $originalString.Substring(6, 2)
    $hour = $originalString.Substring(8, 2)
    $minute = $originalString.Substring(10, 2)
    $second = $originalString.Substring(12, 2)
    $millisecond = $originalString.Substring(15, 3)
    $offset = $originalString.Substring(18)  # 这里假设偏移量为固定的长度，可以根据实际情况进行调整

    # 格式化日期时间字符串
    $formattedDateTime = "{0}年{1}月{2}日 {3}:{4}:{5}" -f $year, $month, $day, $hour, $minute, $second
    
    return $formattedDateTime

}

function get_ad_services( $servername ){
    $ad_services = @()
    try{
        $Active_Directory_Domain_Services = Get-Service -ComputerName $servername -DisplayName "Active Directory Domain Services"
    }catch{
        $Active_Directory_Domain_Services = [PSCustomObject]@{
                    name =  "NTDS"
                    DisplayName = "Active Directory Domain Services"
                    Status = "error"
        }
    }

    try{
        $Active_Directory_Web_Services = Get-Service -ComputerName $servername -DisplayName "Active Directory Web Services"
    }catch{
        $Active_Directory_Domain_Services = [PSCustomObject]@{
                    name =  "ADWS"
                    DisplayName = "Active Directory Web Services"
                    Status = "error"
        }
    }

    try{
        $DNS_Server = Get-Service -ComputerName $servername -DisplayName "DNS Server"
    }catch{
        $DNS_Server = [PSCustomObject]@{
                    name =  "DNS"
                    DisplayName = "DNS Server"
                    Status = "error"
        } 
    }

    try{
        $Kerberos_Key_Distribution_Center = Get-Service -ComputerName $servername -DisplayName "Kerberos Key Distribution Center"
    }catch{
        $Kerberos_Key_Distribution_Center = [PSCustomObject]@{
                    name =  "Kdc"
                    DisplayName = "Kerberos Key Distribution Center"
                    Status = "error"
        } 
    }

    try{
        $Windows_Time = Get-Service -ComputerName $servername -DisplayName "Windows Time"
    }catch{
        $Windows_Time =  [PSCustomObject]@{
                    name =  "W32Time"
                    DisplayName = "Windows Time"
                    Status = "error"
        } 
    }

    try{
        $Netlogon = Get-Service -ComputerName $servername -DisplayName "Netlogon"
    }catch{
        $Netlogon =  [PSCustomObject]@{
                    name =  "Netlogon"
                    DisplayName = "Netlogon"
                    Status = "error"
        } 
    }

    try{
        $DFS_Replication = Get-Service -ComputerName $servername -DisplayName "DFS Replication"
    }catch{
        $DFS_Replication =  [PSCustomObject]@{
                    name =  "DFSR"
                    DisplayName = "DFS Replication"
                    Status = "error"
        } 
    }

    try{
        $Intersite_Messaging = Get-Service -ComputerName $servername -DisplayName "Intersite Messaging"
    }catch{
        $Intersite_Messaging =  [PSCustomObject]@{
                    name =  "IsmServ"
                    DisplayName = "Intersite Messaging"
                    Status = "error"
        } 
    }


    try{
        $Security_Accounts_Manager = Get-Service -ComputerName $servername -DisplayName "Security Accounts Manager"
    }catch{
        $Security_Accounts_Manager =  [PSCustomObject]@{
                    name =  "SamSs"
                    DisplayName = "Security Accounts Manager"
                    Status = "error"
        } 
    }

    $ad_services += $Active_Directory_Domain_Services
    $ad_services += $Active_Directory_Web_Services
    $ad_services += $DNS_Server
    $ad_services += $Kerberos_Key_Distribution_Center
    $ad_services += $Windows_Time
    $ad_services += $Netlogon
    $ad_services += $DFS_Replication
    $ad_services += $Intersite_Messaging
    $ad_services += $Security_Accounts_Manager

    return $ad_services
    <#
        $ad_services =  [PSCustomObject]@{
                    name =  $servername
                    Active_Directory_Domain_Services =  $Active_Directory_Domain_Services
                    Active_Directory_Web_Services =  $Active_Directory_Web_Services
                    DNS_Server =  $DNS_Server
                    Kerberos_Key_Distribution_Center =  $Kerberos_Key_Distribution_Center
                    Windows_Time =  $Windows_Time
                    Netlogon =  $Netlogon
                    DFS_Replication =  $DFS_Replication
                    Intersite_Messaging =  $Intersite_Messaging
                    Security_Accounts_Manager =  $Security_Accounts_Manager
                                
    }
    #>
}

function dcdiag_check($server){
    $info = @()
    $info = dcdiag /s:$server
    $result=@()
    for($i -eq 0;$i -lt $info.count;$i++){
        if($info[$i] -match '开始测试:'){
            $result+=[PSCustomObject]@{
            'ServerName'=$server
            'Item'= $info[$i].replace('开始','')
            'Result'= $info[$i+2].replace('.........................','')
            }
        }
    }
    return $result
}

$time_table = @"
<table>
    <tr>
        <th>服务器</th>
        <th>时间</th>
    </tr>
</table>
"@

foreach ( $dc in $all_dc ){
    $servername = $dc.name
    $time = GetOsTime $servername
    $new_time_row = "
        <tr>
            <td>$servername</td>
            <td>$time</td>
        </tr>"
    $end_time_table_index = $time_table.IndexOf("</table>")
    $time_table = $time_table.Insert($end_time_table_index, $new_time_row)
}

$repl_table = @"
<table>
    <tr>
        <th>目标 DSA</th>
        <th>命名上下文</th>
        <th>源 DSA</th>
        <th>失败次数</th>
        <th>上次成功时间</th>
    </tr>
</table>
"@

$repl_result = @()
$repl_result += repadmin /showrepl ad /csv |  ConvertFrom-Csv | select "目标 DSA","命名上下文","源 DSA","失败次数","上次成功时间"

foreach ( $repl_item in $repl_result ){
   $repl_item_info_01 = $repl_item."目标 DSA"
   $repl_item_info_02 = $repl_item."命名上下文"
   $repl_item_info_03 = $repl_item."源 DSA"
   $repl_item_info_04 = $repl_item."失败次数"
   $repl_item_info_05 = $repl_item."上次成功时间"
   

    $new_repl_row = "
        <tr>
            <td>$repl_item_info_01</td>
            <td>$repl_item_info_02</td>
            <td>$repl_item_info_03</td>
            <td>$repl_item_info_04</td>
            <td>$repl_item_info_05</td>
        </tr>"
    $end_repl_table_index = $repl_table.IndexOf("</table>")
    $repl_table = $repl_table.Insert($end_repl_table_index, $new_repl_row)
}

<#
$repl_result_good_info = @()
foreach ( $repl_item in $repl_result ){
    #$repl_servername = $repl_item.split(" ")[1]
    if ( $repl_item -match "复制" ){
        continue
    }
    if ( $repl_item -match "\.\.\." ){
       continue
    }
    if ( $repl_item -match "目标 DSA" ){
        continue
    }
    $repl_result_good_info += $repl_item
}
$repl_result_from = $repl_result_good_info |  ConvertFrom-Csv
$repl_result_from 
#>


$service_table = @"
<table>
    <tr>
        <th>服务器</th>
        <th>NTDS</th>
        <th>ADWS</th>
        <th>DNS</th>
        <th>Kdc</th>
        <th>W32Time</th>
        <th>Netlogon</th>
        <th>DFSR</th>
        <th>IsmServ</th>
        <th>SamSs</th>
    </tr>
</table>
"@

foreach ( $dc in $all_dc ){
    $servername = $dc.name
    $ad_service = get_ad_services $servername
    $NTDS_status = $ad_service[0].Status
    $ADWS_status = $ad_service[1].Status
    $DNS_status = $ad_service[2].Status
    $Kdc_status = $ad_service[3].Status
    $w32time_status = $ad_service[4].Status
    $Netlogon_status = $ad_service[5].Status
    $DFSR_status = $ad_service[6].Status
    $IsmServ_status = $ad_service[7].Status
    $ISamSs_status = $ad_service[8].Status

    $new_service_row = "
        <tr>
            <td>$servername</td>
            <td>$NTDS_status</td>
            <td>$ADWS_status</td>
            <td>$DNS_status</td>
            <td>$Kdc_status</td>
            <td>$w32time_status</td>
            <td>$Netlogon_status</td>
            <td>$DFSR_status</td>
            <td>$IsmServ_status</td>
            <td>$ISamSs_status</td>
        </tr>"
    $end_service_table_index = $service_table.IndexOf("</table>")
    $service_table = $service_table.Insert($end_service_table_index, $new_service_row)
    
}


$dcdaig_table = @"
<table>
    <tr>
        <th>服务器</th>
        <th>测试项</th>
        <th>结果</th>
    </tr>
</table>
"@
$dcdiag_check_result = @()
$dcdiag_check_result = dcdiag_check ad

foreach ( $check_item in $dcdiag_check_result ){
    $servername = $check_item.ServerName
    $item = $check_item.Item
    $result = $check_item.Result

    $new_dcdiag_row = "
        <tr>
            <td>$servername</td>
            <td>$item</td>
            <td>$result</td>
        </tr>"
    $end_dcdaig_table_index = $dcdaig_table.IndexOf("</table>")
    $dcdaig_table = $dcdaig_table.Insert($end_dcdaig_table_index, $new_dcdiag_row)

}

$body = @"
<html>
<head>
    <style>
        p{height:25px;width:100%;text-align:center;margin:0;font-weight:bold;}
        h1, h5, th { text-align: left; }
        table { margin: auto; font-family: Segoe UI; box-shadow: 10px 10px 5px #888; border: thin ridge grey; }
        th { background: #0046c3; color: #fff; min-width: 120px; padding: 5px 10px; }
        td { font-size: 11px; padding: 5px 20px; color: #000; }
        tr { background: #b8d1f3; }
        tr:nth-child(even) { background: #dae5f4; }
        tr:nth-child(odd) { background: #b8d1f3; }
    </style>
</head>
<body>
    $time_table

    $repl_table

    $service_table

    $dcdaig_table

</body>
</html>
"@



function send_report ( $touser ){
    
    #$UserName = "8000"
    #$Password = "j3HW775S"
    #$Password = ConvertTo-SecureString $Password -AsPlainText -Force
    #$Cred = New-Object System.Management.Automation.PSCredential($UserNAME,$Password)
    $From = "8000@lexin.com"
    $To = $touser + "@lexin.com"
    $subject = 'AD状态报告'

    #$SMTPServer = "mail.lexin.com"
    #$SMTPPort = "587"
    

    #Send-MailMessage -From $From -to $To -Subject $Subject -Body $body -SmtpServer $SMTPServer -port $SMTPPort -Credential $cred -BodyAsHtml -Encoding utf8
    $SMTPServer = "10.1.48.240"
    $SMTPPort = "25"
    Send-MailMessage -From $From -to $To -Subject $Subject -Body $body -SmtpServer $SMTPServer -port $SMTPPort -BodyAsHtml -Encoding utf8
}

send_report junjiesitu
