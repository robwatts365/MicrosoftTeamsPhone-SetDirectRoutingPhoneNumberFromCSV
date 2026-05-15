<# 
Set Direct Routing Phone Numbers and Voice Routing Policies from CSV
    Version: v1.0
    Date: 15/05/2026
    Author: Rob Watts | Sr UCC Engineer - AVI-SPL
#>

#################################################################
################### START OF SCRIPT FUNCTIONS ###################

### FUNCTION - Write To Log File ###
function Write-LogFileMessage($message) {
    $ActionTime = $(Get-Date).ToString("yyyy/MM/dd HH:mm:ss")
    Write-Debug "$ActionTime :: $message"
    "$ActionTime :: $message" | Out-File -FilePath $LogFile -Append
}

################### END OF SCRIPT FUNCTIONS ###################
###############################################################

###############################################################
################### START OF SCRIPT ###########################

# Display Header, Version, Date, Author
Write-Host "---------------------------------" -ForegroundColor Gray
Write-Host "Microsoft Teams: Set Direct Routing Phone Numbers and Voice Routing Policies from CSV" -ForegroundColor Gray -BackgroundColor Black
Write-Host "Version: v1.0" -ForegroundColor Gray -BackgroundColor Black
Write-Host "Date: 15/05/2026" -ForegroundColor Gray -BackgroundColor Black
Write-Host "Author: Rob Watts | Sr UCC Engineer - AVI-SPL" -ForegroundColor Gray -BackgroundColor Black

# Enable File Saver for Log File
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

# Set Log File Location
Write-Host "---------------------------------" -ForegroundColor Gray 
Write-Host "Please select the Log File location."
$LogSaver = New-Object -Typename System.Windows.Forms.SaveFileDialog
$LogSaver.initialDirectory = $initialDirectory
$LogSaver.filter = "All files (*.log)| *.log"
$LogSaver.ShowDialog() | Out-Null
$LogFile = $LogSaver.filename

# Log Header, Version, Date, Author
Write-LogFileMessage "---------------------------------"
Write-LogFileMessage "Microsoft Teams: Set Direct Routing Phone Numbers and Voice Routing Policies from CSV"
Write-LogFileMessage "Version: v1.0"
Write-LogFileMessage "Date: 15/05/2026"
Write-LogFileMessage "Author: Rob Watts | Sr UCC Engineer - AVI-SPL"

# Checks if MicrosoftTeams Powershell Module is installed
Write-Host "---------------------------------" -ForegroundColor Gray
Write-LogFileMessage "---------------------------------"
Write-Host "Checking for Required PowerShell Module..." -ForegroundColor Gray -BackgroundColor Black
Write-LogFileMessage "Checking for Required PowerShell Module..."

# MicrosoftTeams
If (-not(Get-InstalledModule MicrosoftTeams -ErrorAction silentlycontinue)) {
    Write-Host "MicrosoftTeams PowerShell module does not exist... Please Install MicrosoftTeams PowerShell Module..." -ForegroundColor DarkRed
    Write-LogFileMessage "MicrosoftTeams PowerShell module does not exist... Please Install MicrosoftTeams PowerShell Module..."
    Exit
}
Else {
    Write-Host "MicrosoftTeams PowerShell module exists..." -ForegroundColor Green
    Write-LogFileMessage "MicrosoftTeams PowerShell module exists..."
}

# Import Teams Module
Write-Host "---------------------------------" -ForegroundColor Gray
Write-LogFileMessage "---------------------------------"
Write-Host "Importing Microsoft Teams PowerShell Module..." -ForegroundColor Gray -BackgroundColor Black
Write-LogFileMessage "Importing Microsoft Teams PowerShell Module..."    
Import-Module MicrosoftTeams

# Connects to Microsoft Teams
Write-Host "---------------------------------" -ForegroundColor Gray
Write-LogFileMessage "---------------------------------"
Write-Host "Connecting to Microsoft Teams..." -ForegroundColor Gray -BackgroundColor Black
Write-LogFileMessage "Connecting to Microsoft Teams..."
Connect-MicrosoftTeams
Write-Host "---------------------------------" -ForegroundColor Gray
Write-LogFileMessage "---------------------------------"
Write-Host "Connected to Microsoft Teams..."
Write-LogFileMessage "Connected to Microsoft Teams..."

# Enable File Picker
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
Write-Host "---------------------------------" -ForegroundColor Gray
Write-LogFileMessage "---------------------------------"
Write-Host "Please select the CSV file containing User details. Look for a pop out window if not visible." -ForegroundColor Gray -BackgroundColor Black
Write-LogFileMessage "Please select the CSV file containing User details. Look for a pop out window if not visible."

# File Picker  (Set File Path - Open File Browser)
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.initialDirectory = $initialDirectory
$OpenFileDialog.filter = "All files (*.csv)| *.csv"
$OpenFileDialog.ShowDialog() | Out-Null
$FilePath = $OpenFileDialog.filename
   
# Store the data from CSV file in the $Users variable
Write-Host "---------------------------------" -ForegroundColor Gray
Write-LogFileMessage "---------------------------------"
Write-Host "Importing CSV..."
Write-LogFileMessage "Importing CSV from $FilePath"
$Users = Import-Csv $FilePath

# Define Voice Routing Policy to assign to users
$VoiceRoutingPolicy = Get-CsOnlineVoiceRoutingPolicy | Select-Object Identity, Description, OnlinePstnUsages | Out-GridView -OutputMode Single -Title "Please select a Voice Routing Policy to assign to the users"
Write-Host $VoiceRoutingPolicy.Identity "is your chosen Voice Routing Policy"
Write-LogFileMessage "$($VoiceRoutingPolicy.Identity) is your chosen Voice Routing Policy"

# Loop through each row containing user details in the CSV file
foreach ($User in $Users) {

    # Read user data from each field in each row and assign the data to a variable as below
    $UPN = $User.UPN
    $TelephoneNumber = $User.TelephoneNumber
    
    #Assign the Direct Routing Phone Number to the user
    try {
        $SetPhoneNumber = Set-CsPhoneNumberAssignment -Identity $UPN -PhoneNumber $TelephoneNumber -PhoneNumberType DirectRouting -ErrorAction SilentlyContinue -ErrorVariable SetPhoneNumberError
                    
        # Check if phone number assignment returned any output
        if ($null -eq $SetPhoneNumber -or ($SetPhoneNumber | Measure-Object).Count -eq 0) {
            Write-Host "Phone number '$TelephoneNumber' assigned to '$UPN' (no output returned)." -ForegroundColor Green
            Write-LogFileMessage "Phone number '$TelephoneNumber' assigned to '$UPN' (no output returned)."
        }
        else {
            Write-Host "Phone number '$TelephoneNumber' assigned to '$UPN' successfully." -ForegroundColor Green
            Write-LogFileMessage "Phone number '$TelephoneNumber' assigned to '$UPN' successfully."
        }
    }
    catch {
        Write-Host "ERROR: Failed to assign phone number '$TelephoneNumber' to '$UPN'. Error: $SetPhoneNumberError.Record" -ForegroundColor DarkRed
        Write-LogFileMessage "ERROR: Failed to assign phone number '$TelephoneNumber' to '$UPN'. Error: $SetPhoneNumberError.Record"
        continue
    }

    #Assign the Voice Routing Policy to the user
    try {
        Grant-CsOnlineVoiceRoutingPolicy -Identity $UPN -PolicyName $VoiceRoutingPolicy.Identity -ErrorAction SilentlyContinue -ErrorVariable GrantVoiceRoutingPolicyError

        # Check if voice routing policy assignment returned any output
        if ($null -eq $GrantVoiceRoutingPolicy -or ($GrantVoiceRoutingPolicy | Measure-Object).Count -eq 0) {
            Write-Host "Voice Routing Policy '$($VoiceRoutingPolicy.Identity)' assigned to '$UPN' (no output returned)." -ForegroundColor Green
            Write-LogFileMessage "Voice Routing Policy '$($VoiceRoutingPolicy.Identity)' assigned to '$UPN' (no output returned)."
        }
        else {
            Write-Host "Voice Routing Policy '$($VoiceRoutingPolicy.Identity)' assigned to '$UPN' successfully." -ForegroundColor Green
            Write-LogFileMessage "Voice Routing Policy '$($VoiceRoutingPolicy.Identity)' assigned to '$UPN' successfully."
        }
    }
    catch {
        Write-Host "ERROR: Failed to assign Voice Routing Policy '$($VoiceRoutingPolicy.Identity)' to '$UPN'. Error: $GrantVoiceRoutingPolicyError.Record" -ForegroundColor DarkRed
        Write-LogFileMessage "ERROR: Failed to assign Voice Routing Policy '$($VoiceRoutingPolicy.Identity)' to '$UPN'. Error: $GrantVoiceRoutingPolicyError.Record"
        continue
    }

}
Write-Host "---------------------------------" -ForegroundColor Gray
Write-LogFileMessage "---------------------------------"
Write-Host "Script Completed!" -ForegroundColor Green
Write-LogFileMessage "Script Completed!"
pause

################### END OF SCRIPT #############################
###############################################################