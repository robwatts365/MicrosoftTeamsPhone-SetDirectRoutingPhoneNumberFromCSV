<# 
Set Direct Routing Phone Numbers from CSV
    Version: v1.0
    Date: 08/05/2024
    Author: Rob Watts https://github.com/robwatts365
#>

# Import Teams Module
Import-Module MicrosoftTeams

# Connects to Microsoft Teams
Write-Host "Connecting to Microsoft Teams..."
Connect-MicrosoftTeams

# Enable File Picker
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

# File Picker  (Set File Path - Open File Browser)
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "All files (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $FilePath = $OpenFileDialog.filename
   
# Store the data from NewUsersFinal.csv in the $ADUsers variable
    Write-Host "Importing CSV..."
    $Users = Import-Csv $FilePath

# Loop through each row containing user details in the CSV file
foreach ($User in $Users) {

    # Read user data from each field in each row and assign the data to a variable as below
    $UPN = $User.UPN
    $TelephoneNumber = $User.TelephoneNumber
        
    Set-CSPhoneNumberAssignment -Identity $UPN -PhoneNumber $TelephoneNumber -PhoneNumberType DirectRouting

    }
