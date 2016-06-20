param([String]$csvFile="bamboohr.csv")
$office365session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $cred -Authentication Basic -AllowRedirection
Import-PSSession $office365session
$csv = Import-CSV $csvFile
$csv | ForEach-Object{
    $ID = $_."Work Email"
    $LastFirst = $_."Last Name, First Name"
    $WorkPhone = $_."Work Phone"
    $Office = $_.Location
    $City = $Office.split(',')[0]
    $State = $Office.split(',')[1].TrimStart()
    $Department = $_.Department
    $JobTitle = $_."Job Title"
    $split = $LastFirst.split(',')
    $Last = $split[0]
    $First = $split[1].TrimStart()
    $Initials = $First.toCharArray()[0] + $Last.toCharArray()[0]
    $DisplayName = $First + " " + $Last
    $ManagerName = $_."Reporting to"
    $ManagerSplit = $ManagerName.split(" ")
    $ManagerID = $ManagerSplit[0].toCharArray()[0] + $ManagerSplit[1] + "@microscan.com"
    Write-Host "ID: " $ID
    Write-Host "Work Phone: " $WorkPhone
    Write-Host "Office:" $Office
    Write-Host "City: " $City
    Write-Host "State: " $State
    Write-Host "Department: " $Department
    Write-Host "Jobtitle: " $JobTitle
    Write-Host "First Name: " $First
    Write-Host "Last Name: " $Last
    Write-Host "Initials: " $Initials
    Write-Host "Display Name: " $DisplayName
    Write-Host "Manager ID : " $ManagerID
    Set-User -Identity $ID  -City $office.split(',')[0] -Company "Microscan" -Department $Department -DisplayName $DisplayName -FirstName $First -Initials $Initials -LastName $Last -Manager $ManagerID -Office $Office -Phone $WorkPhone -StateOrProvince $State -Title $JobTitle -Verbose
    sleep 1
}
Write-Host "Press any key to continue ..."
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")