param([String]$csvFile="bamboohr.csv")
#setup CSV Field Constants
$ID_COL_NAME = "Work Email"
$LAST_FIRST_COL_NAME = "Last Name, First Name"
$NICKNAME_COL_NAME = "Nickname"
$WORK_PHONE_COL_NAME = "Work Phone"
$MOBILE_PHONE_COL_NAME = "Mobile Phone"
$LOCATION_COL_NAME = "Location"
$DEPARTMENT_COL_NAME = "Department"
$JOB_TITLE_COL_NAME = "Job Title"
$MANAGER_NAME_COL_NAME = "Reporting to"
$COMPANY = "Microscan"

#set up our office 365 session (will prompt for credientials for now)
<#
TODO: Set up service account and assign this script to it permentantly, maybe make credentials read out of a file?
 #>
$office365session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $cred -Authentication Basic -AllowRedirection
Import-PSSession $office365session

#Import our CSV file
$csv = Import-CSV $csvFile

#This map is used to resolve Supervisor names (Non-Nickname) to their email (sometimes starts with first letter of nickname, not real name)
$EmailMap = @{}
#Populate the map
$csv | ForEach-Object{
    $LastFirst = $_.$LAST_FIRST_COL_NAME
    $split = $LastFirst.split(',')
    $Last = $split[0]
    $First = $split[1].TrimStart()
    $First_Last = $First + " " + $Last
    $EmailMap.Add($First_Last,$_."Work Email")
    Write-Host $First_Last " = " $EmailMap.$First_Last
}

#Parses user data from the CSV and then finally sets it in exchange online
$csv | ForEach-Object{
    $ID = $_.$ID_COL_NAME
    $LastFirst = $_.$LAST_FIRST_COL_NAME
    $split = $LastFirst.split(',')
    $First = ""
    if($_.$NICKNAME_COL_NAME -eq ""){
        $First = $split[1].TrimStart()
    }else{
        $First = $_.$NICKNAME_COL_NAME
    }
    $Last = $split[0]
    $WorkPhone = $_.$WORK_PHONE_COL_NAME
    $CellPhone = $_.$MOBILE_PHONE_COL_NAME
    $Office = $_.$LOCATION_COL_NAME
    $City = ""
    $State = ""
    if($Office -like '*,*'){
            $City = $Office.split(',')[0]
            $State = $Office.split(',')[1].TrimStart()
        }else{
            $City = $Office
        }
    $Department = $_.$DEPARTMENT_COL_NAME
    $JobTitle = $_.$JOB_TITLE_COL_NAME
    $ManagerName = $_.$MANAGER_NAME_COL_NAME
    $ManagerEmail = $EmailMap.$ManagerName
    Write-Host "ID:         " $ID
    Write-Host "Work Phone: " $WorkPhone
    Write-Host "Cell Phone: " $CellPhone
    Write-Host "Office:     " $Office
    Write-Host "City:       " $City
    Write-Host "State:      " $State
    Write-Host "Department: " $Department
    Write-Host "Jobtitle:   " $JobTitle
    Write-Host "First Name: " $First
    Write-Host "Last Name:  " $Last
    Write-Host "Manager:    " $ManagerName 
    Write-Host "Manager ID: " $ManagerEmail
    Set-User -Identity $ID  -City $City -Company $COMPANY -Department $Department -FirstName $First -LastName $Last -Manager $ManagerEmail -MobilePhone $CellPhone -Office $Office -Phone $WorkPhone -StateOrProvince $State -Title $JobTitle -Verbose
    sleep 1
}
Write-Host "Press any key to continue ..."
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")