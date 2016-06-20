param([String]$csvFile="bamboohr.csv")
$office365session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $cred -Authentication Basic -AllowRedirection
Import-PSSession $office365session
$csv = Import-CSV $csvFile
$csv | ForEach-Object{
    $ID = $_."Work Email"
    $LastFirst = $_."Last Name, First Name"
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
    $split = $LastFirst.split(',')
    $Last = $split[0]
    $First = $split[1].TrimStart()
    $DisplayName = $First + " " + $Last
    $ManagerName = $_."Reporting to"
    $ManagerSplit = $ManagerName.split(" ")
    $ManagerID = ""
    if($ManagerSplit.Length -eq 2){
        $ManagerID = ($ManagerSplit[0].toCharArray()[0] + ($ManagerSplit[1] -replace "'", "")  + "@microscan.com") -replace ".com.cn" , ".com"
    }else{
        $ManagerID = $ManagerSplit[0].toCharArray()[0]
        for($i = 1; $i -le $ManagerSplit.length; $i++){
            $ManagerID = $ManagerID + $ManagerSplit[$i]
        }
        $ManagerID = $ManagerID + "@microscan.com"
    }
    Write-Host "ID: " $ID
    Write-Host "Work Phone: " $WorkPhone
    Write-Host "Office:" $Office
    Write-Host "City: " $City
    Write-Host "State: " $State
    Write-Host "Department: " $Department
    Write-Host "Jobtitle: " $JobTitle
    Write-Host "First Name: " $First
    Write-Host "Last Name: " $Last
    Write-Host "Display Name: " $DisplayName
    Write-Host "Manager ID : " $ManagerID
    Set-User -Identity $ID  -City $City -Company "Microscan" -Department $Department -FirstName $First -LastName $Last -Manager $ManagerID -Office $Office -Phone $WorkPhone -StateOrProvince $State -Title $JobTitle -Verbose
    sleep 1
}
Write-Host "Press any key to continue ..."
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")