import-module ActiveDirectory
Add-Type -AssemblyName PresentationCore, PresentationFramework
$logfilepath=".\Log.txt"
if(Test-Path $logfilepath)
{
    "~-~-~-~-~-~-~-~-~-~-~-~-~- New Run -~-~-~-~-~-~-~-~-~-~-~-~-~-~" >> $logfilepath
}

#function to see if the requested user exists in AD 
function Test-ADUser {
    param(
        [Parameter(Mandatory = $true)]
        [String] $sAMAccountName
    )

    try {
        Get-ADUser -Identity $sAMAccountName
        $UserExists = $true
    }
    catch {
        $UserExists = $false
    }
    return $UserExists
}

Function GenerateUsers {
    
    param(
        [Parameter(Mandatory = $true)]
        [String] $newUserFirstName,
        [Parameter(Mandatory = $true)]
        [String] $newUserLastName,
        [Parameter(Mandatory = $true)]
        [String] $newUserEmail,
        [Parameter(Mandatory = $true)]
        [String] $newUserDisplayName,
        [Parameter(Mandatory = $true)]
        [String] $newUserUserName,
        [Parameter(Mandatory = $true)]
        [String] $oldUser
    )
    #retrieve info from oldUser such as description, department, member of,  security... etc 
    $user = Get-ADUser $oldUser -Properties Department, Description, Manager, MemberOf, Office, Organization, ProfilePath, Title, Company

    #create new user with firstName, lastName, userName, email and everything else

    New-ADUser -Name $newUserUserName -UserPrincipalName $newUserUserName -DisplayName $newUserDisplayName -AccountPassword (ConvertTo-SecureString -AsPlainText "Password1" -force) -ChangePasswordAtLogon $true -GivenName $newUserFirstName -Surname $newUserLastName -EmailAddress $newUserEmail -Instance $user

    
    #Copy Groups over
    $d = Get-ADPrincipalGroupMembership -Identity $oldUser | Select-Object Name
    Foreach ($g IN $d) {
        if ($g.name -ne 'Domain Users') {
            try {
                Add-ADGroupMember -Identity $g.name -Members $newUserUserName
            }
            catch {
                $counter += 1
            }
        }
    }
    #change new user OU location

    $UserDN = (Get-ADUser -Identity $oldUser).distinguishedName

    $TargetOU = $UserDN.Substring($UserDN.IndexOf('OU='))
    $UserDN2 = (Get-ADUser -Identity $newUserUserName).distinguishedName

    Move-ADObject  -Identity $UserDN2 -TargetPath $TargetOU 

}


#code here for import excel sheet
Write-Host ("")
Write-Host ("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
Write-Host("Welcome to the mass User Creation Tool!")
Write-Host ("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
Write-Host("CSV files should be formatted as follows: ")
$ExampleTable = @(
    @{firstname="joe";lastname="diaz";email="jdiaz@example.com";displayname="diaz, joe";username="jdiaz";},
    @{firstname="joe2";lastname="diaz2";email="jdiaz@example.com2";displayname="diaz2, joe2";username="jdiaz2";}
    
) 
$ExampleTable | ForEach-Object {[PSCustomObject]$_} | Format-Table -Property firstname, lastname, email, displayname, username

Write-Host("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
Write-Host("speficy a local path to a excel file to import it.")
#get path
$path = Read-Host "Enter the path here in the format 'C:\Example\exampleExcel.csv' "
#import csv at path
$users = Import-Csv -Path $path | Select-Object -Property firstname, lastname, email, displayname, username, oldusername

Write-Host(" ")
$oldUserUserName = Read-Host "Enter the username of the user you would like to copy from "

#wait until the inputted old user is a valid user that exists
while ((Test-ADUser($oldUserUserName)) -eq $false){
    [System.Windows.MessageBox]::Show('User copying from does not exist')
    $oldUserUserName = Read-Host "Enter the username of the user to copy from again"
}
$i = 0
Write-Progress -Activity "Creating" -Status "Progress:" -PercentComplete 1
foreach ($user in $users){
    $i = $i+1
    Write-Host ("Creating User with credentials :"+$user)
    # handle any last names that contain a hyphon by taking the first letter
    # of the first hyphonated section and the second hyphonated section concatinated
    # together as the last name
    # ~~
    # lastname: 'test-joe' would be come 'tjoe'
    if ($user.lastname.Contains('-')){
        $nameArray = $user.lastname.split("-")
        for ($x=0; $x -lt $nameArray.count-1;$x++){
            $newLastName += $nameArray[$x].Substring(0,1)    
        }
        $newLastName += $nameArray[$nameArray.count-1]
        $user.lastname = $newLastName
    }
    if ((Test-ADUser($user.username)) -eq $true) {
        #user with that username already exists
        $e = 1
        $newUserNameString = $user.username
        #while username is taken, append more letters of the first name to the username until it is not taken.
        #if newusernamestring is null it is because an error was thrown and so the while loop will be exited and user will not be created.
        while (($null -ne $newUserNameString) -and ((Test-ADUser($newUserNameString)) -eq $true)){
            $e = $e + 1
            #if run out of letters to use in first name
            if ($e -gt $user.firstname.Length){
                #highly unlikely throw error
                $errorMessage = "user "+$user.username+" must be manually created. All letters of first name attempted to use unsuccessful."
                Write-Error -Message $errorMessage
                $newUserNameString = $null
                #generate string and append to log file
                "User "+$user.username +" could not be created because all letters of first name were attempted use. Manually create user with middle name initial."+ (Get-Date).ToString() >> $logfilepath
            }
            # if all is well with adding letters from users first name
            else {
                #add letter from first name and loop again
                $modifiedFirstName = $user.firstname.Substring(0,$e)
                $newUserNameString = $modifiedFirstName + $user.lastname
                Write-Host ("User "+ $user.username +" already exists and was not created. Attempting with "+$newUserNameString)
            }
        }

        if ($null -ne $newUserNameString){
            $newEmail = $newUserNameString + "@prhc.on.ca"
            Write-Host ("User "+ $newUserNameString +" Generating")
            GenerateUsers $user.firstname $user.lastname $newEmail $user.displayname $newUserNameString $oldUserUserName
            #generate string and append to log file
            $userString = "User Created with credentials firstname: "+$user.firstname+", lastname: "+ $user.lastname +", email: "+$newEmail +", displayname: "+$user.displayname+", username: "+ $newUserNameString+", refrenceUser: "+ $oldUserUserName
            $userString +" - "+ (Get-Date).ToString() >> $logfilepath
            Write-Progress -Activity "Creating" -Status "Progress:" -PercentComplete ($i/$users.count*100)
        }
    }
    else {
        #create new user
        Write-Host ("User "+ $user.username +" Generating")
        GenerateUsers $user.firstname $user.lastname $user.email $user.displayname $user.username $oldUserUserName
        #generate string and append to log file
        $userString = "firstname: "+$user.firstname+", lastname: "+ $user.lastname +", email: "+$user.email +", displayname: "+$user.displayname+", username: "+ $user.username+", refrenceUser: "+ $oldUserUserName
        "User Created with credentials "+$userString +" - "+ (Get-Date).ToString() >> $logfilepath

        Write-Progress -Activity "Creating" -Status "Progress:" -PercentComplete ($i/$users.count*100)
    }
    Write-Host " "
}
[System.Windows.MessageBox]::Show('Done.')
Read-Host


