# SyncroScripts - Collection of SyncroMSP scripts 

## Sync folder
Contains scripts to make API calls between different providers like Microsoft 365, Passportal etc
Please see individual scripts for how to use them. 

## MessageAsSystem
It bugged me for a while why Syncro can only send a message when you run scripts as a Logged in user. Most of the scripts required SYSTEM permissions, and there was no way to provide feedback to the user when a script was finished.
Simple solution was to run a check for any user with explorer.exe running (desktop experience) and set up scheduled tasks to create notifications for these users. Just like that – we can notify any user without worrying about if we run as a user or not.
I’m hoping that Syncro crew will be able to get that pushed into one of the updates within module.psm1 ( more info at https://bit.ly/2S8U4JI ) 
In the meantime, if you need to send message to users when running scripts as SYSTEM – here is function that you could place on the top of your script

## Reset Local Admin Password
Durring our onboarding process we create local admin user with random password containing our special prefix and X number of random words from dictionary. 
This script is set to weekly reset them and upload new password back to Syncro. 

### Contributions
Feel free to send pull requests or fill out issues when you encounter them. I'm also completely open to adding maintainers/contributors and working together! :)
