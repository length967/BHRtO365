param([String]$csvFile="bamboohr.csv")
$office365session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $cred -Authentication Basic -AllowRedirection
Import-PSSession $office365session
$csv = Import-CSV $csvFile

$EmailMap = @{}

$csv | ForEach-Object{
    $LastFirst = $_."Last Name, First Name"
    $split = $LastFirst.split(',')
    $Last = $split[0]
    $First = $split[1].TrimStart()
    $First_Last = $First + " " + $Last
    $EmailMap.Add($First_Last,$_."Work Email")
    Write-Host $First_Last " = " $EmailMap.$First_Last
}

$csv | ForEach-Object{
    $ID = $_."Work Email"
    $LastFirst = $_."Last Name, First Name"
    $split = $LastFirst.split(',')
    $First = ""
    if($_.Nickname -eq ""){
        $First = $split[1].TrimStart()
    }else{
        $First = $_.Nickname
    }
    $Last = $split[0]
    $WorkPhone = $_."Work Phone"
    $Office = $_.Location
    $City = ""
    $State = ""
    if($Office -like '*,*'){
            $City = $Office.split(',')[0]
            $State = $Office.split(',')[1].TrimStart()
        }else{
            $City = $Office
        }
    $Department = $_.Department
    $JobTitle = $_."Job Title"
    $ManagerName = $_."Reporting to"
    $ManagerEmail = $EmailMap.$ManagerName
    Write-Host "ID: " $ID
    Write-Host "Work Phone: " $WorkPhone
    Write-Host "Office:" $Office
    Write-Host "City: " $City
    Write-Host "State: " $State
    Write-Host "Department: " $Department
    Write-Host "Jobtitle: " $JobTitle
    Write-Host "First Name: " $First
    Write-Host "Last Name: " $Last
    Write-Host "Manager: " $ManagerName 
    Write-Host "Manager ID : " $ManagerEmail
    Set-User -Identity $ID  -City $City -Company "Microscan" -Department $Department -FirstName $First -LastName $Last -Manager $ManagerEmail -Office $Office -Phone $WorkPhone -StateOrProvince $State -Title $JobTitle -Verbose
    sleep 1
}
Write-Host "Press any key to continue ..."
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")