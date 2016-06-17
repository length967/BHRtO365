param([String]$csvFile="bamboohr.csv")
$office365session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $cred -Authentication Basic -AllowRedirection
Import-PSSession $office365session
$csv = Import-CSV $csvFile
$csv | ForEach-Object{
    $ID = $_."Work Email"
    $LastFirst = $_."Last Name, First Name"
    $WorkPhone = $_."Work Phone"
    $Office = $_.Location
    $Department = $_.Department
    $JobTitle = $_."Job Title"
    $split = $LastFirst.split(',')
    $Last = $split[0]
    $First = $split[1].TrimStart()
    $DisplayName = $First + " " + $Last
    Write-Host "ID: " $ID
    Write-Host "Work Phone: " $WorkPhone
    Write-Host "Office:" $Office
    Write-Host "Department: " $Department
    Write-Host "Jobtitle: " $JobTitle
    Write-Host "First Name: " $First
    Write-Host "Last Name: " $Last
    Write-Host "Display Name: " $DisplayName
}
Write-Host "Press any key to continue ..."
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")