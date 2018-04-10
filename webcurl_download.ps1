$day=(Get-date).ToString("MMddyyyy.'csv'")
D:
cd D:\ServerEPS_active
del Export TRX01*.csv
curl -k "https://www.sitename.com/servervice/export/alltransactions.xpt?companynumber=XXXX&username=xxxxxxxx&password=xxxxxx&format=CSV" > "Export TRX01 $day"