#These scripts below are merged into one to show a process flow.
#These scripts are not the best way to do what it is designed to do (export data from multiple sources, merge and parse it)  
#This could be done much more efficiently importing all external data sources to DB tables and doing the work there.
#I simply use this "Rube-Goldberg" methodology to display different functionality in one process.


cd c:\export

del *.xlsx
del *.csv

Set-ExecutionPolicy AllSigned
#Connection Strings

$Database = "xxxxxxxx"
$Start= ((Get-Date).AddDays(-7)).ToString("MM-dd-yyyy")
$End= ((Get-Date).AddDays(-1)).ToString("MM-dd-yyyy")
$Server = "xxxxxxxxx"
$AttachmentPath = "C:\PR.csv"

# Clear Hash Table

clear $objTable_PR
clear $objTable_SLs
clear $objTable_Regular

#Create HQ data files

\\xxxxxxxxx\instore\hqm\hqm /user xxxxxxx /password xxx /menu XXXXXXXXwithaddedfields410 "SendTo File=c:\HQ_410.csv, Close" /Exit | Out-Null
\\xxxxxxxxx\instore\hqm\hqm /user xxxxxxx /password xxx /menu XXXXXXXXwithaddedfields411 "SendTo File=c:\HQ_411.csv, Close" /Exit | Out-Null
\\xxxxxxxxx\instore\hqm\hqm /user xxxxxxx /password xxx /menu XXXXXXXXwithaddedfields417 "SendTo File=c:\HQ_417.csv, Close" /Exit | Out-Null

# Connect to SQL and query data, extract data to temp table, pare down temp table.
$SqlQuery = "SELECT (CASE WHEN b.store_number IN (02, 04, 05, 06, 15, 18, 12, 27, 28, 29) THEN 410 WHEN b.store_number IN (10, 14) 
                      THEN 411 WHEN b.store_number IN (11, 12, 13, 16, 17, 19) THEN 417 END) AS ZN, a.CP, SUM(b.RTL_dollars) AS SLs, SUM(b.TGE) AS TGE, SUM(b.QTTY) AS QTTY
Into #PR
FROM       xxxxxxxx.dbo.item_master AS a INNER JOIN
           xxxxxxxx.dbo.item_xxxxxxxx AS b ON a.item_id = b.item_id
WHERE     (a.minor_department IN ('01', '02', '03', '04', '05', '06', '08', '09', '11', '12', '13', '14', '15')) AND (b.xxxxxxxx_date BETWEEN '05/04/2012' AND '05/10/2012') AND (a.item_status <> 'Discontinued') AND 
                      (b.store_number <> 'CMPNY') AND (a.minor_category IN ('001100', '00102', '00104', '00105', '00106', '001107', '002101', '003100', '003101', '010110', '01101', 
                      '014100'))
and xxxxxxxx_ID IN (1,2,4)
					  GROUP BY store_number, CP
					  
select ZN, CP, 
Sum (SLs) as PR_SLs,
Sum (TGE) as PR_TGE,
Sum (QTTY) as PR_QTTY
from #PR GROUP BY ZN, CP
order by CP"

$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Data Source=$Server;Initial Catalog=$Database;Integrated Security = True"
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlCmd.CommandText = $SqlQuery
$SqlCmd.Connection = $SqlConnection
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCmd
$DataSet = New-Object System.Data.DataSet
$nRecs = $SqlAdapter.Fill($DataSet)
$nRecs | Out-Null

#Populate Hash Table

$objTable_PR = $DataSet.Tables[0]

#Export Hash Table to CSV File

$objTable_PR | Export-CSV -NoTypeInformation $AttachmentPath

#create date variable

$date = Get-Date -format "yyyyMMdd"

#create/work with Hash Tables

$grouped_PR_Test1_410=Import-Csv c:\PR.csv |WHERE-OBJECT {$_.ZN -eq 410}|  group CP -AsHashTable -AsString 
Import-Csv c:\HQ_410.csv  | foreach{
	$SLs=($grouped_PR_Test1_410."$($_.CP)" | foreach {$_.PR_SLs}) -join ","
	 $_ | Add-Member -MemberType NoteProperty -Name SLS -Value $SLs -PassThru 
 } | Export-Csv -NoTypeInformation c:\HQ_PR_SLs_410.csv

$grouped_PR_Test1_411=Import-Csv c:\PR.csv |WHERE-OBJECT {$_.ZN -eq 411}|  group CP -AsHashTable -AsString 
Import-Csv c:\HQ_411.csv  | foreach{
	$SLs=($grouped_PR_Test1_411."$($_.CP)" | foreach {$_.PR_SLs}) -join ","
	 $_ | Add-Member -MemberType NoteProperty -Name SLS -Value $SLs -PassThru 
 } | Export-Csv -NoTypeInformation c:\HQ_PR_SLs_411.csv

$grouped_PR_Test1_417=Import-Csv c:\PR.csv |WHERE-OBJECT {$_.ZN -eq 417}|  group CP -AsHashTable -AsString 
Import-Csv c:\HQ_417.csv  | foreach{
	$SLs=($grouped_PR_Test1_417."$($_.CP)" | foreach {$_.PR_SLs}) -join ","
	 $_ | Add-Member -MemberType NoteProperty -Name SLS -Value $SLs -PassThru 
 } | Export-Csv -NoTypeInformation c:\HQ_PR_SLs_417.csv
 
$grouped_PR_Test1_410.Clear() 
$grouped_PR_Test1_411.Clear() 
$grouped_PR_Test1_417.Clear() 

#del HQ_4*.csv

$grouped_PR_Test2_410=Import-Csv c:\PR.csv |WHERE-OBJECT {$_.ZN -eq 410}|  group CP -AsHashTable -AsString
Import-Csv c:\HQ_PR_SLs_410.csv | foreach{
	$TGE=($grouped_PR_Test2_410."$($_.CP)" | foreach {$_.PR_TGE}) -join ","
	 $_ | Add-Member -MemberType NoteProperty -Name TGE -Value $TGE -PassThru 
 } | Export-Csv -NoTypeInformation c:\HQ_PR_SLs_TGE_410.csv
 
 $grouped_PR_Test2_411=Import-Csv c:\PR.csv |WHERE-OBJECT {$_.ZN -eq 411}|  group CP -AsHashTable -AsString
Import-Csv c:\HQ_PR_SLs_411.csv | foreach{
	$TGE=($grouped_PR_Test2_411."$($_.CP)" | foreach {$_.PR_TGE}) -join ","
	 $_ | Add-Member -MemberType NoteProperty -Name TGE -Value $TGE -PassThru 
 } | Export-Csv -NoTypeInformation c:\HQ_PR_SLs_TGE_411.csv
 
 $grouped_PR_Test2_417=Import-Csv c:\PR.csv |WHERE-OBJECT {$_.ZN -eq 417}|  group CP -AsHashTable -AsString
Import-Csv c:\HQ_PR_SLs_417.csv | foreach{
	$TGE=($grouped_PR_Test2_417."$($_.CP)" | foreach {$_.PR_TGE}) -join ","
	 $_ | Add-Member -MemberType NoteProperty -Name TGE -Value $TGE -PassThru 
 } | Export-Csv -NoTypeInformation c:\HQ_PR_SLs_TGE_417.csv
 
$grouped_PR_Test2_410.Clear() 
$grouped_PR_Test2_411.Clear() 
$grouped_PR_Test2_417.Clear() 
 
 $grouped_PR_Test3_410=Import-Csv c:\PR.csv |WHERE-OBJECT {$_.ZN -eq 410}|  group CP -AsHashTable -AsString
Import-Csv c:\HQ_PR_SLs_TGE_410.csv | foreach{
	$QTTY=($grouped_PR_Test3_410."$($_.CP)" | foreach {$_.PR_QTTY}) -join ","
	 $_ | Add-Member -MemberType NoteProperty -Name QTTY -Value $QTTY -PassThru 
 } | Export-Csv -NoTypeInformation "c:\XXXXXXXX_05_410_1.csv"
 
 $grouped_PR_Test3_411=Import-Csv c:\PR.csv |WHERE-OBJECT {$_.ZN -eq 411}|  group CP -AsHashTable -AsString
Import-Csv c:\HQ_PR_SLs_TGE_411.csv | foreach{
	$QTTY=($grouped_PR_Test3_411."$($_.CP)" | foreach {$_.PR_QTTY}) -join ","
	 $_ | Add-Member -MemberType NoteProperty -Name QTTY -Value $QTTY -PassThru 
 } | Export-Csv -NoTypeInformation "c:\XXXXXXXX_14_411_1.csv"
 
 $grouped_PR_Test3_417=Import-Csv c:\PR.csv |WHERE-OBJECT {$_.ZN -eq 417}|  group CP -AsHashTable -AsString
Import-Csv c:\HQ_PR_SLs_TGE_417.csv | foreach{
	$QTTY=($grouped_PR_Test3_417."$($_.CP)" | foreach {$_.PR_QTTY}) -join ","
	 $_ | Add-Member -MemberType NoteProperty -Name QTTY -Value $QTTY -PassThru 
 } | Export-Csv -NoTypeInformation "c:\XXXXXXXX_17_417_1.csv"
 
$grouped_PR_Test3_410.Clear() 
$grouped_PR_Test3_411.Clear() 
$grouped_PR_Test3_417.Clear() 
 
 #Split columns for Ct_N / sbCt_N and  Cat_ID / sbCt_ID
 
 import-csv c:\XXXXXXXX_05_410_1.csv | ForEach-Object {
    $_.Sub_Ct_N,$tempCAT1=$_.Sub_Ct_N -split "-",2
    $_.Sb_Cat_ID,$tempID1=$_.Sb_Cat_ID -split " ",2
    $_ | Select-Object  CP, Pdct_Dec, Sz, UM, @{ expression={$_.Sub_Ct_N}; label='Ct_N' }, @{Name="Sub_Ct_N";Expression={$tempCAT1}},  @{ expression={($_.Sb_Cat_ID).Substring(0,3)}; label='Cat_ID'}, @{ expression={($_.Sb_Cat_ID).Substring(3,3)}; label='Sb_Cat_ID'}, Vdr, Vdr_ID, "Reg Pr Mult", Base_Prc, "SL Pr Mult", Prm_Prc, Promo_Strt_Dt, Prm_nd_Dt, "PR Pr Mult", PR_Prc, PR_Strt_Dt, PR_Ed_Dt, ZN_Prc_Strg, Prc_ASSC_CD, IP_Prc_MTD, SL_IP_RTL_DSCN, SL_RTL_DSC_FLG, PR_IP_RTL_DSCN, PR_RTL_DSC_FLG, Dp., "Dp Desc", SLS, TGE, QTTY
} | export-csv  -NoTypeInformation c:\XXXXXXXX_05_410.csv

import-csv c:\XXXXXXXX_14_411_1.csv | ForEach-Object {
    $_.Sub_Ct_N,$tempCAT2=$_.Sub_Ct_N -split "-",2
    $_.Sb_Cat_ID,$tempID2=$_.Sb_Cat_ID -split " ",2
    $_ | Select-Object  CP, Pdct_Dec, Sz, UM, @{ expression={$_.Sub_Ct_N}; label='Ct_N' }, @{Name="Sub_Ct_N";Expression={$tempCAT2}},  @{ expression={($_.Sb_Cat_ID).Substring(0,3)}; label='Cat_ID'}, @{ expression={($_.Sb_Cat_ID).Substring(3,3)}; label='Sb_Cat_ID'}, Vdr, Vdr_ID, "Reg Pr Mult", Base_Prc, "SL Pr Mult", Prm_Prc, Promo_Strt_Dt, Prm_nd_Dt, "PR Pr Mult", PR_Prc, PR_Strt_Dt, PR_Ed_Dt, ZN_Prc_Strg, Prc_ASSC_CD, IP_Prc_MTD, SL_IP_RTL_DSCN, SL_RTL_DSC_FLG, PR_IP_RTL_DSCN, PR_RTL_DSC_FLG, Dp., "Dp Desc", SLS, TGE, QTTY
} | export-csv  -NoTypeInformation c:\XXXXXXXX_14_411.csv

import-csv c:\XXXXXXXX_17_417_1.csv | ForEach-Object {
    $_.Sub_Ct_N,$tempCAT3=$_.Sub_Ct_N -split "-",2
    $_.Sb_Cat_ID,$tempID3=$_.Sb_Cat_ID -split " ",2
    $_ | Select-Object  CP, Pdct_Dec, Sz, UM, @{ expression={$_.Sub_Ct_N}; label='Ct_N' }, @{Name="Sub_Ct_N";Expression={$tempCAT3}},  @{ expression={($_.Sb_Cat_ID).Substring(0,3)}; label='Cat_ID'}, @{ expression={($_.Sb_Cat_ID).Substring(3,3)}; label='Sb_Cat_ID'}, Vdr, Vdr_ID, "Reg Pr Mult", Base_Prc, "SL Pr Mult", Prm_Prc, Promo_Strt_Dt, Prm_nd_Dt, "PR Pr Mult", PR_Prc, PR_Strt_Dt, PR_Ed_Dt, ZN_Prc_Strg, Prc_ASSC_CD, IP_Prc_MTD, SL_IP_RTL_DSCN, SL_RTL_DSC_FLG, PR_IP_RTL_DSCN, PR_RTL_DSC_FLG, Dp., "Dp Desc", SLS, TGE, QTTY
} | export-csv  -NoTypeInformation c:\XXXXXXXX_17_417.csv

exit
