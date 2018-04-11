#download file via URL using cURL 

Set-ExecutionPolicy AllSigned

$day=(Get-date).ToString("MMddyyyy.'csv'")
D:
cd D:\SerPS_act
del Export TR01*.csv
curl -k "https://www.sitename.com/servervice/export/alltransactions.xpt?companynumber=XXXX&username=xxxxxxxx&password=xxxxxx&format=CSV" > "Export TR01 $day"
