$connectSplat = @{
    CertificateFilePath = 'c:\Cert.pfx'
    CertificatePassword = $(ConvertTo-SecureString -String "" -AsPlainText -Force)
    AppID = ''
    Organization = ""

}

Connect-ExchangeOnline @connectSplat

$domains = @('@client.com','@client.com','@client.COM','@client.nyc')

$start_date = (Get-Date).AddDays(-1)
$start_date =  $start_date.ToString("MM/dd/yyyy") + ' 12:00:00 AM'
$end_date = (Get-Date).AddDays(-1)
$end_date =  $end_date.ToString("MM/dd/yyyy") + ' 11:59:00 PM'
$title = Get-Date -Format "MM.dd.yyyy"
$table = New-Object System.Data.Datatable
[void]$table.Columns.Add("Name")
[void]$table.Columns.Add("Count")

foreach ($domain in $domains) {

    $All_count = Get-MessageTrace -SenderAddress *@source4.com -recipient *$domain -start $start_date -end $end_date | Group-Object count
    $Ex_count = Get-MessageTrace -SenderAddress ExcaliburWMS@source4.com -recipient *$domain -start $start_date -end $end_date | Group-Object count
    $total = $All_count.Count - $Ex_count.Count
    $total = $total.ToString()
    if ($domain.Contains("wildone")) {$domain = "WILDONE"}
    ElseIf ($domain.Contains("staycourant")) {$domain = "COURANT"}
    ElseIf ($domain.Contains("WANDPDESIGN")) {$domain = "WPDESIGN"}
    ElseIf ($domain.Contains("verygreat")) {$domain = "VERYGREAT"}
    [void]$table.Rows.Add("$domain","$total")

}

$table | export-csv "file.csv" -NoTypeInformation

Disconnect-ExchangeOnline -Confirm:$false
