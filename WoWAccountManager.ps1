########################
##WoW Account Manager##
########################
function InitializeAccounts{get-childitem "$env:USERPROFILE\Documents\WindowsPowerShell\Scripts\encrypted" -Filter "*.txt" -ErrorAction SilentlyContinue}
function MainMenu {
###################
##Begin Main Menu##
###################
## The user is prompted to make some encrypted credentials if there are none found.
## If there are encrypted credentials found, the user is prompted to choose an account to log in with.
Write-Host "Welcome to the WoW Auto-Login Script" -ForegroundColor Yellow -BackgroundColor Black
Write-Host "Added accounts can be deleted here: $env:USERPROFILE\Documents\WindowsPowerShell\Scripts\encrypted" -ForegroundColor Yellow -BackgroundColor Black

# Check if any encrypted cached accounts were found
$getAccounts = InitializeAccounts
if ($getAccounts) {
    Write-Host "$($getAccounts.Count) cached account(s) was found" -ForegroundColor Yellow
    foreach($account in $getAccounts){ Write-Host "- $($account.BaseName) `n"}
    WoWHandler
} else {
    Write-Host "No cached accounts were found... Loading Credential Handler." -ForegroundColor Yellow
    CredentialHandler
}
}

Function FindWoW {

##################
##Locate WoW.exe##
##################
#Searches the drive you choose for the WoW.exe, it first finds the WoW folder and then locates the .exe inside.


# Define the target file name you're searching for
$fileName = "WoW.exe"
# Define the folder name pattern that includes wildcards to match versioned folders (e.g., "World of Warcraft 1.12")
$folderName = "World of Warcraft"

# Get all available file system drives (including local drives)
$drives = Get-WmiObject Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 }

# Prompt the user with a dropdown to choose a drive
$selectedDrive = $drives | Select-Object -ExpandProperty DeviceID | Out-GridView -Title "Select the drive where WoW is installed" -PassThru

# Check if a drive was selected
if ($selectedDrive) {

Write-Host "Drive $selectedDrive was selected. Please wait while WoW.exe is located." -ForegroundColor Yellow


# Get list of directories under $selectedDrive (limit depth to 3 levels for optimization)
$folders = Get-ChildItem "$selectedDrive\" -Recurse -Directory -Depth 3 -ErrorAction SilentlyContinue

# Filter only directories named "World of Warcraft"
$wowFolders = $folders | Where-Object { $_.Name -like "$folderName *" }

# Total number of folders found
$totalFolders = $wowFolders.Count
$currentFolder = 0

# Output the number of directories found (optional for debug)
Write-Output "$totalFolders 'World of Warcraft' folders found."

# Loop through the "World of Warcraft" folders and search for WoW.exe
$foundFile = $null
foreach ($folder in $wowFolders) {
    $currentFolder++

    # Update progress bar
    Write-Progress -Activity "Searching for WoW.exe..." -Status "Processing Folder $currentFolder of $totalFolders" -PercentComplete (($currentFolder / $totalFolders) * 100)

    # Search for WoW.exe in the current "World of Warcraft" folder
    $file = Get-ChildItem -Path $folder.FullName -Recurse -Filter $fileName -File -ErrorAction SilentlyContinue

    # If WoW.exe is found, store it and break the loop
    if ($file) {
        $foundFile = $file
        break
    }

    # Small delay for smooth progress bar updates
    Start-Sleep -Milliseconds 50
}

# Clear the progress bar after completion
Write-Progress -Activity "Completed" -Status "Done" -Completed

# Output the found file if it was found, or notify that it wasn't found
if ($foundFile) {
    Write-Host "Found file: $($foundFile.FullName) `n Saving location to $env:USERPROFILE\Documents\WindowsPowerShell\Scripts\WoWLocation.txt." -ForegroundColor Yellow
        
    # Define the path for the file
    $WoWSavedfilePath = "$env:USERPROFILE\Documents\WindowsPowerShell\Scripts\WoWLocation.txt"

        # Check if the directory exists, and if not, create it
        $directory = [System.IO.Path]::GetDirectoryName($WoWSavedfilePath)
        if (-not (Test-Path $directory)) {
        New-Item -ItemType Directory -Path $directory -Force
        }
    
    # Save the WoW file path to a text file after making sure it exists.
    $foundFile.FullName | Out-File "$env:USERPROFILE\Documents\WindowsPowerShell\Scripts\WoWLocation.txt" -Force
} else {
    Write-Output "WoW.exe not found."
}
}
}


############################
##Begin Credential Handler##
############################

Function CredentialHandler {

# Check if the directory exists, and if not, create it
$EncryptedCredentialDirectory = "$env:USERPROFILE\Documents\WindowsPowerShell\Scripts\encrypted"
if (-not (Test-Path $EncryptedCredentialDirectory)) {
    $null = New-Item -ItemType Directory -Path $EncryptedCredentialDirectory -Force
}

$addAccount = Read-Host "Do you want to add a new encrypted account file? (Y/N)"
$addAccount = $addAccount.ToUpper()
if ($addAccount -eq "N") {MainMenu}
While ($addAccount -eq "Y") {

# Prompt the user for credentials
$credential = Get-Credential -Message "Enter the WoW Account credentials to be saved to an encrypted file."

# Convert the password to a secure string and save it to a file
$credential.Password | ConvertFrom-SecureString | Out-File "$env:USERPROFILE\Documents\WindowsPowerShell\Scripts\encrypted\$($credential.UserName).txt"

Write-Host "Credentials for $($credential.UserName) have been saved." -ForegroundColor Green
Write-Host "Would you like to add another account?" -ForegroundColor Yellow
CredentialHandler

}
MainMenu

}

#####################
##Begin WoW Handler##
#####################

function WoWHandler {

# choose the account you want to log in with
$account = Get-ChildItem "$env:USERPROFILE\Documents\WindowsPowerShell\Scripts\encrypted" | Out-GridView -Title "Select the account you want to log in with" -PassThru
if (-not $account) {
    Write-Host "No account selected. Exiting script." -ForegroundColor Red
    exit
}
$AccountEncrypted = Get-Content "$env:USERPROFILE\Documents\WindowsPowerShell\Scripts\encrypted\$account" | ConvertTo-SecureString
$credential = New-Object System.Management.Automation.PSCredential ($account.BaseName, $AccountEncrypted)
# Launch Application
Start-Process "$wowLocation" -PassThru

# Wait for application to launch
Start-Sleep -Seconds 5

# Send keys to application
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.SendKeys]::SendWait($account.BaseName)
[System.Windows.Forms.SendKeys]::SendWait("{TAB}")
[System.Windows.Forms.SendKeys]::SendWait([Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($credential.Password)))
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")

}


$wowLocation = get-content "$env:USERPROFILE\Documents\WindowsPowerShell\Scripts\WoWLocation.txt" -ErrorAction SilentlyContinue
if($wowLocation){Write-Host "WoW.exe location is saved here: `n$wowLocation" -ForegroundColor Yellow}
# This is the function (FindWoW) that will build out the directory we need to store the encrypted credentials and save the location of the WoW.exe file if this is the first time running the script.
else{FindWoW}

MainMenu