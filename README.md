# World-of-Warcraft Utilities

# World of Warcraft Account Manager Utility


I was tired of copy and pasting account info out of a password manager and decided to make a WoW Account Manager program in Powershell. This is usefull for players who use multiple accounts and want to expedite logins.
- To Do:
    - ~~make it so that after initial setup, you can go back in and add more accounts.~~
        - (done)
    - ~~add accomidations for Vanilla Tweaks and Vanilla Fixes .exe's~~
        - (done)

# Overview of Functions

Initialization and Main Menu:
The script initializes by checking for existing encrypted account files.
It displays a main menu with options to launch WoW, add/remove accounts, or reinitialize the WoW executable location.

Finding WoW Executable:
The FindWoW function searches for the WoW.exe file (or alternatives like VanillaFixes.exe and WoW_Tweaked.exe) on the selected drive.
It saves the path of the found executable to a text file for future use.

Credential Handling:
The CredentialHandler function provides options to add or remove accounts or return to the main menu.
Adding Accounts:

The AddAccount function prompts the user to enter WoW account credentials, which are then encrypted and saved to a file.

Removing Accounts:
The RemoveAccount function lists the saved encrypted account files and allows the user to select and delete one.
WoW Handler:

The WoWHandler function allows the user to select an account to log in with.
It launches the WoW application, waits for it to start, and then sends the username and password to the application using SendKeys.
Special characters in the password are handled by the Convert-SpecialCharacters function to ensure they are sent correctly.
The script is designed to manage WoW account credentials securely and automate the login process. 
