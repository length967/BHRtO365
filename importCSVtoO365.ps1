param([String]$csvFile="bamboohr.csv")
#setup CSV Field Constants
$ID_COL_NAME = "Work Email"
$LAST_FIRST_COL_NAME = "Last Name, First Name"
$NICKNAME_COL_NAME = "Preferred Name"
$WORK_PHONE_COL_NAME = "Work Phone"
$MOBILE_PHONE_COL_NAME = "Mobile Phone"
$LOCATION_COL_NAME = "Location"
$DEPARTMENT_COL_NAME = "Department"
$JOB_TITLE_COL_NAME = "Job Title"
$MANAGER_NAME_COL_NAME = "Reporting to"
$COMPANY = "Microscan"
$US_STATE_NAMES = "Alabama","Alaska","Arizona","Arkansas","California","Colorado","Connecticut","Delaware","Florida","Georgia","Hawaii","Idaho","Illinois","Indiana","Iowa","Kansas","Kentucky","Louisiana","Maine","Maryland","Massachusetts","Michigan","Minnesota","Mississippi","Missouri","Montana","Nebraska","Nevada","New Hampshire","New Jersey","New Mexico","New York","North Carolina","North Dakota","Ohio","Oklahoma","Oregon","Pennsylvania","Rhode Island","South Carolina","South Dakota","Tennessee","Texas","Utah","Vermont","Virginia","Washington","West Virginia","Wisconsin","Wyoming"
$US_STATE_ABBRV =  "AK","AL","AZ","AR","CA","CO","CT","DE","FL","GA","HI","ID","IL","IN","IA","KS","KY","LA","ME","MD","MA","MI","MN","MS","MO","MT","NE","NV","NH","NJ","NM","NY","NC","ND","OH","OK","OR","PA","RI","SC","SD","TN","TX","UT","VT","VA","WA","WV","WI","WY"

#Other Constants
$TPAD = 45

#set up our office 365 session (will prompt for credientials for now)
<#
TODO: Set up service account and assign this script to it permentantly, maybe make credentials read out of a file?
 #>
$cred = Get-Credential
$office365session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $cred -Authentication Basic -AllowRedirection
Import-PSSession $office365session

#Import our CSV file
$csv = Import-CSV $csvFile

#This map is used to resolve Supervisor names (Non-Nickname) to their ID (sometimes starts with first letter of nickname, not real name)
$EmailMap = @{}
#Populate the map
$csv | ForEach-Object{
    $LastFirst = $_.$LAST_FIRST_COL_NAME
    $split = $LastFirst.split(',')
    $Last = $split[0]
    $First = $split[1].TrimStart()
    $First_Last = $First + " " + $Last
    $EmailMap.Add($First_Last,($_."Work Email" -replace "@.*\..*", ""))
    Write-Host $First_Last " = " $EmailMap.$First_Last
}

#Parses user data from the CSV, compare to current user data, then if different, finally sets it in exchange online
$csv | ForEach-Object{
    $ID = $_.$ID_COL_NAME
    $CurrentUser = ""
    Try {
        $CurrentUser = Get-User -Identity $ID -ErrorAction Stop
    }
    Catch{
        Write-Host "User: " $ID " was not found in O365 Database"
        continue
    }
    $LastFirst = $_.$LAST_FIRST_COL_NAME
    $split = $LastFirst.split(',')
    $First = ""
    $Last = $split[0]
    if($_.$NICKNAME_COL_NAME -eq ""){
        $First = $split[1].TrimStart()
    }else{
        $First = $_.$NICKNAME_COL_NAME
    }
    $WorkPhone = $_.$WORK_PHONE_COL_NAME
    $CellPhone = $_.$MOBILE_PHONE_COL_NAME
    $Office = $_.$LOCATION_COL_NAME
    #US and remote US user detection
    $US_STATE_NAMES | ForEach-Object{
        if($Office -like ("*"+$_+"*")){
            Write-Host $Office, " = ", $_
            Set-Mailbox -Identity $ID -CustomAttribute2 "USUSER"
            if($Office -notlike "*WA" -and $Office -notlike "*NH"){
            Set-Mailbox -Identity $ID -CustomAttribute3 "REMOTE"
            }
        }
    }
    $US_STATE_ABBRV | ForEach-Object{
        if($Office -like ("*",$_) -and $Offce -notlike "*india*" -and $Office -notlike "*Alphen*" -and $Office -notlike "*China*"){
            Write-Host $Office, " = ", $_
            Set-Mailbox -Identity $ID -CustomAttribute2 "USUSER"
            if($Office -notlike "*WA" -and $Office -notlike "*NH"){
            Set-Mailbox -Identity $ID -CustomAttribute3 "REMOTE"
            }
        }
    }
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
    $ManagerID = $EmailMap.$ManagerName
    if(!$ManagerID) {$ManagerID = " "}
    $ManagerName = (get-user -Identity $ManagerID).DisplayName
    if($ManagerName -eq $null) {$ManagerID = " "}
    $CurrentManager = $currentUser.manager.displayName
    if(!$CurrentManager) {$CurrentManager = (get-user -Identity $CurrentUser.manager).DisplayName}
    if($CurrentManager -eq $null) {$CurrentManager = " "}
    $WorkPhoneChanged = $CellPhoneChanged = $OfficeChanged = $CityChanged = $StateChanged = $DepartmentChanged = $JobTitleChanged = $FirstNameChanged = $LastNameChanged = $ManagerChanged = ""
    if($CurrentUser.FirstName -ne $First){
        Set-User -Identity $ID -FirstName $First
        $FirstNameChanged = "X"
    }
    if($CurrentUser.LastName -ne $Last){
        Set-User -Identity $ID -LastName $Last
        $LastNameChanged = "X"
    }
    if($CurrentUser.phone -ne $WorkPhone){
        Set-User -Identity $ID -Phone $WorkPhone
        $WorkPhoneChanged = "X"
    }
    if($CurrentUser.MobilePhone -ne $CellPhone){
        Set-User -Identity $ID -MobilePhone $CellPhone
        $CellPhoneChanged = "X"
    }
    if($CurrentUser.Office -ne $Office){
        Set-User -Identity $ID -Office $Office
        $OfficeChanged = "X"
    }
    if($CurrentUser.City -ne $City){
        Set-User -Identity $ID -City $City
        $CityChanged = "X"
    }
    if($CurrentUser.StateOrProvince -ne $State){
        Set-User -Identity $ID -StateOrProvince $State
        $StateChanged = "X"
    }
    if($CurrentUser.Department -ne $Department){
        Set-User -Identity $ID -Department $Department
        $DepartmentChanged = "X"
    }
    if($CurrentUser.Title -ne $JobTitle){
        Set-User -Identity $ID -Title $JobTitle
        $JobTitleChanged = "X"
    }
    if($CurrentManager -ne $ManagerName){
        Set-User -Identity $ID -Manager $ManagerID
        $JobTitleChanged = "X"
    }
    Set-Mailbox -Identity $ID -CustomAttribute1 "USER"
    Write-Host "Field Name |" "From CSV".padright($TPAD) "|" "Current/Old".padright($TPAD) "|Changed"
    Write-Host "--------------------------------------------------------------------------------------------------------------------"
    Write-Host "ID         |" $ID.padright($TPAD) "|" $CurrentUser.ID.padright($TPAD) "|" 
    Write-Host "First Name |" $First.padright($TPAD) "|" $CurrentUser.FirstName.padright($TPAD) "|" $FirstNameChanged
    Write-Host "Last Name  |" $Last.padright($TPAD) "|" $CurrentUser.LastName.padright($TPAD) "|" $LastNameChanged
    Write-Host "Work Phone |" $WorkPhone.padright($TPAD) "|" $CurrentUser.phone.padright($TPAD) "|" $WorkPhoneChanged
    Write-Host "Cell Phone |" $CellPhone.padright($TPAD) "|" $CurrentUser.MobilePhone.padright($TPAD) "|" $CellPhoneChanged
    Write-Host "Office     |" $Office.padright($TPAD) "|" $CurrentUser.Office.padright($TPAD) "|" $OfficeChanged
    Write-Host "City       |" $City.padright($TPAD) "|" $CurrentUser.City.padright($TPAD) "|" $CityChanged
    Write-Host "State      |" $State.padright($TPAD) "|" $CurrentUser.StateOrProvince.padright($TPAD) "|" $StateChanged
    Write-Host "Department |" $Department.padright($TPAD) "|" $CurrentUser.Department.padright($TPAD) "|" $DepartmentChanged
    Write-Host "JobTitle   |" $JobTitle.padright($TPAD) "|" $CurrentUser.Title.padright($TPAD) "|" $JobTitleChanged
    Write-Host "Manager    |" $ManagerName.padright($TPAD) "|" $CurrentManager.padright($TPAD) "|" $ManagerChanged
    Write-Host ""
    Write-Host ""
   # Set-User -Identity $ID  -City $City -Company $COMPANY -Department $Department -FirstName $First -LastName $Last -Manager $ManagerID -MobilePhone $CellPhone -Office $Office -Phone $WorkPhone -StateOrProvince $State -Title $JobTitle -Verbose
}
Write-Host "Press any key to continue ..."
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")