import-module ActiveDirectory
Add-Type -AssemblyName PresentationCore, PresentationFramework

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
$ExampleTable | ForEach {[PSCustomObject]$_} | Format-Table -Property firstname, lastname, email, displayname, username

Write-Host("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
Write-Host("speficy a local path to a excel file to import it.")
#get path
$path = Read-Host "Enter the path here in the format 'C:\Example\exampleExcel.csv' "
#import csv at path
$users = Import-Csv -Path $path | Select-Object -Property firstname, lastname, email, displayname, username, oldusername

Write-Host(" ")
$oldUserUserName = Read-Host "Enter the username of the user you would like to copy from "

while ((Test-ADUser($oldUserUserName)) -eq $false){
    [System.Windows.MessageBox]::Show('User copying from does not exist')
    $oldUserUserName = Read-Host "Enter the username of the user to copy from again"
}
$i = 0
Write-Progress -Activity "Creating" -Status "Progress:" -PercentComplete 1
foreach ($user in $users){
    $i = $i+1
    Write-Host ("Creating User with credentials :"+$user)
    if ((Test-ADUser($user.username)) -eq $true) {
        Write-Progress -Activity "Creating" -Status "Progress:" -PercentComplete ($i/$users.count*100)
        Write-Host ("User "+ $user.username +" already exists and was not created.")
    }
    else {
        
        Write-Host ("User "+ $user.username +" Generating")
        
        GenerateUsers $user.firstname $user.lastname $user.email $user.displayname $user.username $oldUserUserName
        Write-Progress -Activity "Creating" -Status "Progress:" -PercentComplete ($i/$users.count*100)
    }
    
}
[System.Windows.MessageBox]::Show('Done.')
Read-Host


