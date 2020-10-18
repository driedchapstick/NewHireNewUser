clear

#--------------------------------------------------------------------------------------
#Functions needed to adjust data from OSTicket's MySQL Database.
#--------------------------------------------------------------------------------------

#If no reformatting is needed, just make sure the string is trimmed up.
function TrimMeUp{
    PARAM([String]$providedAnswer)
    
    IF($providedAnswer -eq $null){
        #Do Nothing
    }ELSE{
        $providedAnswer = $providedAnswer.Trim()
    }

    return $providedAnswer
}

#The Dates Need to Be Reformatted
function FixAnswersFromTime{
    PARAM([String]$providedTime)

    $indexOfSpace = $providedTime.IndexOf(' ')
    $providedTime = $providedTime.Substring(0, $indexOfSpace)

    return $providedTime.Trim()
}

#Answers from dropdown lists are wacky
function FixAnswersFromList{
    PARAM([String]$providedAnswer)

    $providedAnswer = $providedAnswer.Replace('{', '')
    $providedAnswer = $providedAnswer.Replace('"', '')
    $providedAnswer = $providedAnswer.Replace('}', '')
    $indexOfColon = $providedAnswer.IndexOf(':')
    $providedAnswer = $providedAnswer.Substring($indexOfColon+1)

    return $providedAnswer.Trim()
}

#Checking for spaces makes sure the user put in first AND last name
function CheckForSpaces{
    PARAM([String]$providedString)

    $spaceInString = $providedString.IndexOf(' ')

    IF($spaceInString -eq -1){
        
        Write-Host ''
        Write-Host 'ERROR: No Space in name.'
        Write-Host '-----------------------'
        Write-Host 'The user provided an improper input.'
        Write-Host "Input`: $providedString"
        Write-Host '-----------------------'
        Write-Host "Check the user's input and properly format it."
        Write-Host ''
        $redoInput = Read-Host 'Properly Formatted Input'

        CheckForSpaces -providedString $redoInput  
    }ELSE{
        return $providedString.Trim()
    }
}

#Format the name to be compliant with the Get-ADUser CMDlet --> Lastname, Firstname
function FormatName{
    PARAM([String]$providedName)

    $spaceInName = $providedName.IndexOf(' ')

    $firName = $providedName.SubString(0, $spaceInName)
    $lasName = $providedName.Substring($spaceInName+1)

    $properName = "$lasName, $firName"

    return $properName.Trim()
}

#Try to find a user account that matches what the user provided. 
function RetrieveSupervisor{
    PARAM([String]$providedSupervisor)

    $originalInput = $providedSupervisor

    $spacedInput = CheckForSpaces -providedString $originalInput
    $formattedInput = FormatName -providedName $spacedInput

    $retrievedUser = ((Get-ADUser -Filter 'name -Like $formattedInput') | Format-Wide -Property name)

    IF(($retrievedUser -eq ' ') -OR ($retrievedUser -eq $null)){
        
        Write-Host ''
        Write-Host 'ERROR: No User Found.'
        Write-Host '-----------------------'
        Write-Host 'Could not find a user with the retrieved name.'
        Write-Host "Name`: $formattedInput"
        Write-Host '-----------------------'
        Write-Host "Please Try Again."
        Write-Host ''
        $redoSupervisor = Read-Host 'Supervisor Name'

        RetrieveSupervisor -providedSupervisor $redoSupervisor
    }ELSE{
        #Need to convert the name format from what is understood by the Get-ADUser cmdlet to what is understood by the New-ADuser
        $managerFirstName = $spacedInput.Substring(0,1)

        $indexOfSpaceForManager = $spacedInput.IndexOf(' ')
        $managerLastName = $spacedInput.Substring($indexOfSpaceForManager+1) 

        $superForNewADUser = $managerFirstName+$managerLastName

        return $superForNewADUser.Trim()
    }
}

#--------------------------------------------------------------------------------------
#Querying OSTicket's MySQL database.
#--------------------------------------------------------------------------------------

Write-Host 'Ticket Number for New Hire'   
$ticketNum = Read-Host 'Include Zeros'

<# Place commands to authenticate and connect to MySQL server.
   Requires the installation of MySQL Connector/Net.
   Requires the installation of the MySQL PowerShell module.#> 

$reqDate= (Invoke-MySqlQuery "Select value FROM ost_ticket JOIN ost_form_entry ON ost_ticket.ticket_id = ost_form_entry.object_id JOIN ost_form_entry_values ON ost_form_entry.id = ost_form_entry_values.entry_id WHERE ost_ticket.number = $ticketNum AND ost_form_entry.form_id = 7 AND ost_form_entry_values.field_id = 36;").Item("value")
$startDate = (Invoke-MySqlQuery "Select value FROM ost_ticket JOIN ost_form_entry ON ost_ticket.ticket_id = ost_form_entry.object_id JOIN ost_form_entry_values ON ost_form_entry.id = ost_form_entry_values.entry_id WHERE ost_ticket.number = $ticketNum AND ost_form_entry.form_id = 7 AND ost_form_entry_values.field_id = 37;").Item("value")
$division = (Invoke-MySqlQuery "Select value FROM ost_ticket JOIN ost_form_entry ON ost_ticket.ticket_id = ost_form_entry.object_id JOIN ost_form_entry_values ON ost_form_entry.id = ost_form_entry_values.entry_id WHERE ost_ticket.number = $ticketNum AND ost_form_entry.form_id = 7 AND ost_form_entry_values.field_id = 38;").Item("value")
$department = (Invoke-MySqlQuery "Select value FROM ost_ticket JOIN ost_form_entry ON ost_ticket.ticket_id = ost_form_entry.object_id JOIN ost_form_entry_values ON ost_form_entry.id = ost_form_entry_values.entry_id WHERE ost_ticket.number = $ticketNum AND ost_form_entry.form_id = 7 AND ost_form_entry_values.field_id = 79;").Item("value")
$supervisor = (Invoke-MySqlQuery "Select value FROM ost_ticket JOIN ost_form_entry ON ost_ticket.ticket_id = ost_form_entry.object_id JOIN ost_form_entry_values ON ost_form_entry.id = ost_form_entry_values.entry_id WHERE ost_ticket.number = $ticketNum AND ost_form_entry.form_id = 7 AND ost_form_entry_values.field_id = 39;").Item("value")
$title = (Invoke-MySqlQuery "Select value FROM ost_ticket JOIN ost_form_entry ON ost_ticket.ticket_id = ost_form_entry.object_id JOIN ost_form_entry_values ON ost_form_entry.id = ost_form_entry_values.entry_id WHERE ost_ticket.number = $ticketNum AND ost_form_entry.form_id = 7 AND ost_form_entry_values.field_id = 40;").Item("value")
$plantLoc = (Invoke-MySqlQuery "Select value FROM ost_ticket JOIN ost_form_entry ON ost_ticket.ticket_id = ost_form_entry.object_id JOIN ost_form_entry_values ON ost_form_entry.id = ost_form_entry_values.entry_id WHERE ost_ticket.number = $ticketNum AND ost_form_entry.form_id = 7 AND ost_form_entry_values.field_id = 66;").Item("value")
$officeLoc = (Invoke-MySqlQuery "Select value FROM ost_ticket JOIN ost_form_entry ON ost_ticket.ticket_id = ost_form_entry.object_id JOIN ost_form_entry_values ON ost_form_entry.id = ost_form_entry_values.entry_id WHERE ost_ticket.number = $ticketNum AND ost_form_entry.form_id = 7 AND ost_form_entry_values.field_id = 41;").Item("value")
$newPosition = (Invoke-MySqlQuery "Select value FROM ost_ticket JOIN ost_form_entry ON ost_ticket.ticket_id = ost_form_entry.object_id JOIN ost_form_entry_values ON ost_form_entry.id = ost_form_entry_values.entry_id WHERE ost_ticket.number = $ticketNum AND ost_form_entry.form_id = 7 AND ost_form_entry_values.field_id = 42;").Item("value")
$replacingUser = (Invoke-MySqlQuery "Select value FROM ost_ticket JOIN ost_form_entry ON ost_ticket.ticket_id = ost_form_entry.object_id JOIN ost_form_entry_values ON ost_form_entry.id = ost_form_entry_values.entry_id WHERE ost_ticket.number = $ticketNum AND ost_form_entry.form_id = 7 AND ost_form_entry_values.field_id = 43;").Item("value")
$firstName = (Invoke-MySqlQuery "Select value FROM ost_ticket JOIN ost_form_entry ON ost_ticket.ticket_id = ost_form_entry.object_id JOIN ost_form_entry_values ON ost_form_entry.id = ost_form_entry_values.entry_id WHERE ost_ticket.number = $ticketNum AND ost_form_entry.form_id = 7 AND ost_form_entry_values.field_id = 44;").Item("value")
$lastName = (Invoke-MySqlQuery "Select value FROM ost_ticket JOIN ost_form_entry ON ost_ticket.ticket_id = ost_form_entry.object_id JOIN ost_form_entry_values ON ost_form_entry.id = ost_form_entry_values.entry_id WHERE ost_ticket.number = $ticketNum AND ost_form_entry.form_id = 7 AND ost_form_entry_values.field_id = 45;").Item("value")
$midName = (Invoke-MySqlQuery "Select value FROM ost_ticket JOIN ost_form_entry ON ost_ticket.ticket_id = ost_form_entry.object_id JOIN ost_form_entry_values ON ost_form_entry.id = ost_form_entry_values.entry_id WHERE ost_ticket.number = $ticketNum AND ost_form_entry.form_id = 7 AND ost_form_entry_values.field_id = 46;").Item("value")
$userName = (Invoke-MySqlQuery "Select value FROM ost_ticket JOIN ost_form_entry ON ost_ticket.ticket_id = ost_form_entry.object_id JOIN ost_form_entry_values ON ost_form_entry.id = ost_form_entry_values.entry_id WHERE ost_ticket.number = $ticketNum AND ost_form_entry.form_id = 7 AND ost_form_entry_values.field_id = 47;").Item("value")
$emailSuffix = (Invoke-MySqlQuery "Select value FROM ost_ticket JOIN ost_form_entry ON ost_ticket.ticket_id = ost_form_entry.object_id JOIN ost_form_entry_values ON ost_form_entry.id = ost_form_entry_values.entry_id WHERE ost_ticket.number = $ticketNum AND ost_form_entry.form_id = 7 AND ost_form_entry_values.field_id = 76;").Item("value")
$computer = (Invoke-MySqlQuery "Select value FROM ost_ticket JOIN ost_form_entry ON ost_ticket.ticket_id = ost_form_entry.object_id JOIN ost_form_entry_values ON ost_form_entry.id = ost_form_entry_values.entry_id WHERE ost_ticket.number = $ticketNum AND ost_form_entry.form_id = 7 AND ost_form_entry_values.field_id = 48;").Item("value")
$computerComm = (Invoke-MySqlQuery "Select value FROM ost_ticket JOIN ost_form_entry ON ost_ticket.ticket_id = ost_form_entry.object_id JOIN ost_form_entry_values ON ost_form_entry.id = ost_form_entry_values.entry_id WHERE ost_ticket.number = $ticketNum AND ost_form_entry.form_id = 7 AND ost_form_entry_values.field_id = 49;").Item("value")
$deskPhone = (Invoke-MySqlQuery "Select value FROM ost_ticket JOIN ost_form_entry ON ost_ticket.ticket_id = ost_form_entry.object_id JOIN ost_form_entry_values ON ost_form_entry.id = ost_form_entry_values.entry_id WHERE ost_ticket.number = $ticketNum AND ost_form_entry.form_id = 7 AND ost_form_entry_values.field_id = 50;").Item("value")

<# Disconnect from the MySQL server. #>

#--------------------------------------------------------------------------------------
#Adjusting the data provided by the query.
#--------------------------------------------------------------------------------------

$reqDate = FixAnswersFromTime -providedTime $reqDate
$startDate = FixAnswersFromTime -providedTime $startDate
$division = FixAnswersFromList -providedAnswer $division
$department = FixAnswersFromList -providedAnswer $department
$supervisor = RetrieveSupervisor -providedSupervisor $supervisor
$manager = $supervisor
$title = TrimMeUp -providedAnswer $title
$plantLoc = FixAnswersFromList -providedAnswer $plantLoc
$officeLoc = TrimMeUp -providedAnswer $officeLoc
$newPosition = TrimMeUp -providedAnswer $newPosition
$replacingUser = TrimMeUp -providedAnswer $replacingUser
$firstName = TrimMeUp -providedAnswer $firstName
$lastName = TrimMeUp -providedAnswer $lastName
$midName = TrimMeUp -providedAnswer $midName
$userName = TrimMeUp -providedAnswer $userName
$emailSuffix = FixAnswersFromList -providedAnswer $emailSuffix
$computer = FixAnswersFromList -providedAnswer $computer
$computerComm = TrimMeUp -providedAnswer $computerComm
$deskPhone = TrimMeUp -providedAnswer $deskPhone

#Determine logon (userPrincipalName) and website (homepage) from Email Suffix
$userPrincipalName = ''
$homepage = ''
$samAccountName = ($firstName.Substring(0,1)+$lastName).ToLower()
$multiEmail = $false
$altEmail = ''
IF(($emailSuffix -eq 'TheArmorGroup.com') -OR ($emailSuffix -eq 'ArmorAftermarket.com') -OR ($emailSuffix -eq 'ArmorConstructionServ.com') -OR ($emailSuffix -eq 'ArmorContract.com') -OR ($emailSuffix -eq 'ArmorProductsInc.com') -OR ($emailSuffix -eq 'CinInd.com') -OR ($emailSuffix -eq 'Witt.com')){
    $userPrincipalName = $samAccountName+'@'+$emailSuffix
    $homepage = "www.$emailSuffix"
}ELSEIF(($emailSuffix -eq 'ArmorMetal.com') -OR ($emailSuffix -eq 'Regency-Recruiting.com')){
    $userPrincipalName = $samAccountName+'@'+$emailSuffix
    $homepage = 'www.TheArmorGroup.com'
}ELSEIF($emailSuffix -eq 'ArmorGov.com'){
    $userPrincipalName = $samAccountName+'@'+$emailSuffix
    $homepage = 'www.ArmorContract.com'
}ELSEIF(($emailSuffix -eq 'ArmorMobie.com') -OR ($emailSuffix -eq 'PQInd.com') -OR ($emailSuffix -eq 'Processall.com')){
    $userPrincipalName = $samAccountName+'@'+'ArmorMetal.com'
    $homepage = "www.$emailSuffix"
    $multiEmail = $true
    $altEmail = $samAccountName+'@'+$emailSuffix
}ELSE{
    Write-Host ''
    Write-Host 'Could not find a corresponding domain for the User Principal Name.'
    Write-Host "Email Suffix that was Parsed: $emailSuffix"
    Write-Host 'The user will be given ArmorMetal.com as a domain and www.TheArmorGroup.com as a webpage.'

    $userPrincipalName = $samAccountName+'@'+'ArmorMetal.com'
    $homepage = 'www.TheArmorGroup.com'
}

#Determine the address from generic plant location
$streetAddress = ''
$city = ''
$state = ''
$postalCode =''
IF($plantLoc -eq "Mason"){
    $streetAddress = '4600 N. Mason-Montgomery Rd'
    $city = 'Mason'
    $state = 'OH'
    $postalCode = '45040'

}ELSEIF($plantLoc -eq "Lebanon"){
    $streetAddress = '160 Harmon Ave'
    $city = 'Lebanon'
    $state = 'OH'
    $postalCode = '45036'

}ELSEIF($plantLoc -eq "Elkhart - Main"){
    $streetAddress = '3362 S Main St'
    $city = 'Elkhart'
    $state = 'IN'
    $postalCode = '46517'

}ELSEIF ($plantLoc -eq "Elkhart - Comet"){
    $streetAddress = '300 Comet Ave'
    $city = 'Elkhart'
    $state = 'IN'
    $postalCode = '456514'

}ELSE{
    Write-Host ''
    Write-Host 'Could not find a corresponding address from the provided Plant Location'
    Write-Host "Plant Location that was Parsed: $plantLoc"
    Write-Host 'The user will be given the Mason Address in their ActiveDirectory Profile.'

    $plantLoc = 'Mason'
    $streetAddress = '4600 N. Mason-Montgomery Rd'
    $city = 'Mason'
    $state = 'OH'
    $postalCode = '45040'
}

#Determine the Exchange Database For the New user.
$exchangeDBNumber = ''
$exchangeDBName = ''

$lastInitial = ($lastName.Substring(0,1)).ToUpper()
IF(($lastInitial -eq 'A') -or ($lastInitial -eq 'B') -or ($lastInitial -eq 'C')){
    $exchangeDBNumber = 0
}ELSEIF(($lastInitial -eq 'D') -or ($lastInitial -eq 'E') -or ($lastInitial -eq 'F') -or ($lastInitial -eq 'G') -or ($lastInitial -eq 'H')){
    $exchangeDBNumber = 1
}ELSEIF(($lastInitial -eq 'I') -or ($lastInitial -eq 'J') -or ($lastInitial -eq 'K') -or ($lastInitial -eq 'L') -or ($lastInitial -eq 'M') -or ($lastInitial -eq 'N')){
    $exchangeDBNumber = 2
}ELSEIF(($lastInitial -eq 'O') -or ($lastInitial -eq 'P') -or ($lastInitial -eq 'Q') -or ($lastInitial -eq 'R') -or ($lastInitial -eq 'S')){
    $exchangeDBNumber = 3
}ELSEIF(($lastInitial -eq 'T') -or ($lastInitial -eq 'U') -or ($lastInitial -eq 'V') -or ($lastInitial -eq 'W') -or ($lastInitial -eq 'X') -or ($lastInitial -eq 'Y') -or ($lastInitial -eq 'Z')){
    $exchangeDBNumber = 4
}ELSE{
    Write-Host ''
    Write-Host 'ERROR: Could not determine the correct database to put the user.'
    Write-Host '----------------------------'
    Write-Host "The user's lastname is $surname."
    Write-Host "The initial used to determine the database is $lastInitial."
    Write-Host '----------------------------'
    DO{
        Write-Host ''
        Write-Host 'Manually Select The Database'
        Write-Host "0. Lastname's Between A-C"
        Write-Host "1. Lastname's Between D-H"
        Write-Host "2. Lastname's Between I-N"
        Write-Host "3. Lastname's Between O-S"
        Write-Host "4. Lastname's Between T-Z"
        Write-Host ''
        Write-Host 'Type the corresponding number.'
        
        $loopAgain = 1
        $firstChoice = Read-Host '0-4'
        Write-Host 'Confirm Choice. (Input Number Again)'
        $secondChoice = Read-Host '0-4'
         IF($firstChoice -eq $secondChoice){
            IF(($firstChoice -eq 0) -or ($firstChoice -eq 1) -or ($firstChoice -eq 2) -or ($firstChoice -eq 3) -or ($firstChoice -eq 4)){
                $loopAgain = 0
                $exchangeDBNumber = $firstChoice
            }ELSE{
                Write-Host ''
                Write-Host 'ERROR: Database Not Found'
                Write-Host '----------------------------'
                Write-Host "You did not provide a valid database ($firstChoice)."
                Write-Host 'Try Again.'
                Write-Host '----------------------------'
            }
        }ELSE{
            Write-Host ''
            Write-Host 'ERROR: Invalid Inputs'
            Write-Host '----------------------------'
            Write-Host "Your first choice ($firstChoice) and second choice ($secondChoice) did not match."
            Write-Host 'Try Again.'
            Write-Host '----------------------------'
        }
    }WHILE($loopAgain -eq 1)
}

IF($exchangeDBNumber -eq 0){
    $exchangeDBName = '*REDACTED*'
}ELSEIF($exchangeDBNumber -eq 1){
    $exchangeDBName = '*REDACTED*'
}ELSEIF($exchangeDBNumber -eq 2){
    $exchangeDBName = '*REDACTED*'
}ELSEIF($exchangeDBNumber -eq 3){
    $exchangeDBName = '*REDACTED*'
}ELSEIF($exchangeDBNumber -eq 4){
    $exchangeDBName = '*REDACTED*'
}

#-------------------------------------------------------------------------------------------------------------------
# Assign values to variables that will ONLY be used in User Creation. 
#-------------------------------------------------------------------------------------------------------------------
$formFirstName = $firstName
$formLastName = $lastName
$formDisplayName = "$lastName, $firstName"
$formSAMAccountName = $samAccountName
$formUserPrincipalName = $userPrincipalName
$formPhoneNumber = '(513) 923-xxxx'
$formHomepage = $homepage
$formCompany = 'The Armor Group Inc.'
$formDepartment = $department
$formOffice = $plantLoc
$formStreetAddress = $streetAddress
$formCity = $city
$formState = $state
$formPostalCode = $postalCode
$formManager = $supervisor
$formTitle = $title
$formMailboxDB = $exchangeDBName
$formRetentionPolicy = 'Armor Default Retention Policy'
$formAltEmail = $altEmail

#-------------------------------------------------------------------------------------------------------------------
# Checks to make sure all the information is correct before continuing.
#-------------------------------------------------------------------------------------------------------------------
Write-Host ''
Write-Host '------------------------------------------------------------'
Write-Host '                   Check the Information                    '
Write-Host '------------------------------------------------------------'
Write-Host "Firstname             -   $formFirstName"
Write-Host "Lastname              -   $formLastName"
Write-Host "Lastname, Firstname   -   $formDisplayName"
Write-Host "FLastname             -   $formSAMAccountName"
Write-Host "FLastName@Domain.com  -   $formUserPrincipalName"
Write-Host "Phone Number          -   $formPhoneNumber"
Write-Host "Homepage              -   $formHomepage"
Write-Host "Company               -   $formCompany"
Write-Host "Department            -   $formDepartment"
Write-Host "Office                -   $formOffice"
Write-Host "Street Address        -   $formStreetAddress"
Write-Host "City                  -   $formCity"
Write-Host "State                 -   $formState"
Write-Host "Postal Code           -   $formPostalCode"
Write-Host "Manager               -   $formManager"
Write-Host "Title                 -   $formTitle"
Write-Host ''
Write-Host "Mailbox Database      -   $formMailboxDB"
Write-Host "Retention Policy      -   $formRetentionPolicy"
Write-Host ''
Write-Host '-----------------------'
Write-Host 'Confirm the Information'
Write-Host '-----------------------'
Write-Host '1 = Everything is Correct.'
Write-Host '0 = Something is Wrong'
Write-Host ''
$isItCorrect = Read-Host 'Answer'

IF($isItCorrect -eq 1){
    
    #-------------------------------------------------------------------------------------------------------------------
    # Creates the new AD Account and enables the Exchange Mailbox
    #-------------------------------------------------------------------------------------------------------------------
    $initalPassword = ConvertTo-SecureString -AsPlainText "*REDACTED*" -Force

    $userAttributes = @{
        GivenName = $formFirstName
        Surname = $formLastName
        Name = $formDisplayName
        DisplayName = $formDisplayName
        SamAccountName = $formSAMAccountName
        UserPrincipalName = $formUserPrincipalName
        OfficePhone = $formPhoneNumber
        Homepage = $formHomepage
        Company = $formCompany
        Department = $formDepartment
        Office = $formOffice
        StreetAddress = $formStreetAddress
        City = $formCity
        State = $formState
        PostalCode = $formPostalCode
        Manager = $formManager
        Title = $formTitle

        AccountPassword = $initalPassword
        Enabled = $true
    }
    
    New-ADUser @userAttributes

    $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri *REDACTED* -Authentication Kerberos
    Import-PSSession $exchangeSession -DisableNameChecking | Out-Null

    #Attempt the mailbox creation in case of small connection issue.
    #Potential Consolidation: Create a function that loops until certain attempts or mailbox is created.
    try{
        Enable-Mailbox -Identity $formSAMAccountName -Alias $formSAMAccountName -RetentionPolicy $formRetentionPolicy -Database $formMailboxDB | Out-Null
        IF($multiEmail -eq $true){
            Set-Mailbox -Identity $formSAMAccountName -EmailAddresses @{add="$formAltEmail"}
        }
    }catch{
        Write-Host ''
        Write-Host 'Failed the First Attempt to Create Exchange Mailbox.'
        try{
            Enable-Mailbox -Identity $formSAMAccountName -Alias $formSAMAccountName -RetentionPolicy $formRetentionPolicy -Database $formMailboxDB | Out-Null
            IF($multiEmail -eq $true){
                Set-Mailbox -Identity $formSAMAccountName -EmailAddresses @{add="$formAltEmail"}
            }
        }catch{
            Write-Host ''
            Write-Host 'Failed the Second Attempt to Create Exchange Mailbox.'
            try{
                Enable-Mailbox -Identity $formSAMAccountName -Alias $formSAMAccountName -RetentionPolicy $formRetentionPolicy -Database $formMailboxDB | Out-Null
                IF($multiEmail -eq $true){
                    Set-Mailbox -Identity $formSAMAccountName -EmailAddresses @{add="$formAltEmail"}
                }
            }catch{
                Write-Host ''
                Write-Host 'Failed the Third Attempt to Create Exchange Mailbox.'
                Write-Host 'The Mailbox will not be created.'

                Write-Host ''
                $error[2]
            }
        }
    }

    Remove-PSSession $exchangeSession
}ELSEIF($isItCorrect -eq 0){
    Write-Host ''
    Write-Host 'The information is not correct.'
    Write-Host 'Edit the supplied information in the New Hire form to make it compatible with the script.'
}ELSE{
    Write-Host ''
    Write-Host 'You did not enter a zero or a one. Run the script again.'
}