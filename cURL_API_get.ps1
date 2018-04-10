#Run cURL connection to vendor to gather path for files to download... results to response_init.xml file

cd c:\curl
.\curl -H "Content-Type: application/xml" -d '@test_post_full.xml' -u username:password -X POST http://api.websitename.com/v1.1/init/xml > response_init.xml
#use contents of response_init.xml to "get" data from host via https, and save.

[xml]$XmlDocument = Get-Content -Path C:\curl\response_init.xml
$XmlDocument.GetType().FullName
$XmlDocument.vendor_data.service.init.response.file_list.file | FORMAT-Table -Wrap -Property file_url  | Out-File c:\curl\URL.txt
    $a = (Get-Content C:\curl\URL.txt)[3]
    $a = $a -replace 'https:','http:'
    .\curl -u username:password -o c:\curl\init.zip "$a"