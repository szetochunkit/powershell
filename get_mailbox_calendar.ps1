# 导入 EWS 程序集
Add-Type -Path "D:\dll\Microsoft.Exchange.WebServices.dll"

# 设置 EWS API 的 URL
$ewsUrl = "https://mail.asdf.com/EWS/Exchange.asmx"

# 创建 Exchange Service 对象
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)

# 设置凭据
$credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials("asdf", "asdasd")
$service.Credentials = $credentials

# 设置 EWS API 的 URL
$service.Url = New-Object System.Uri($ewsUrl)

# 设置查询的时间范围
$startDate = Get-Date "2023-01-01"
$endDate = Get-Date "2023-06-30"

# 设置要查询的邮箱地址
$mailboxAddress = "asdf@asdf.com"

# 创建 FolderId 对象，指定要搜索的邮箱文件夹
$inboxFolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar, $mailboxAddress)

# 创建 CalendarView 对象
$calendarView = New-Object Microsoft.Exchange.WebServices.Data.CalendarView($startDate, $endDate)

# 获取日历项
$calendarItems = $service.FindItems($inboxFolderId, $calendarView)

# 遍历日历项并输出详细信息
foreach ($item in $calendarItems.Items) {
    $item.Load()
    Write-Host "Subject: $($item.Subject)"
    Write-Host "Start Time: $($item.Start)"
    Write-Host "End Time: $($item.End)"
    Write-Host "Location: $($item.Location)"
    Write-Host "------------------------------------"
}
