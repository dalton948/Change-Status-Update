Change Status Update README.


Purpose:
Sends an email to Requestor and all neccessary users for CC updating them when a change has been Canceled, placed on hold, or has failed.

Preliminary Tasks:
1. Run the ACLUPDATEUSERFOLDER.ps1 to update your permissions required for this script to run properly
2. ALTERNATE METHOD - Manually Update your pemissions for C:\users\NON-S USERNAME
	you need to open power shell as an elevated user, 
	type explorer, 
	navigate to C:\Users, 
	right click your non-s folder, 
	click properties, go to security, 
	click edit under the group or user names section, 
	then click Add, 
	you will need to then enter your name as the following example "Dalton Decker S Account", 
	click ok, 
	click Full Control in the permissions for that user you just added, 
	finally click ok and continue through all the prompts that may pop up.

3. Add onedrive shortcute from teams\command\files
	Navigate to Teams,
	go to command channel,
	click on "Add shortcut to OneDrive"


How To Use:
1. Update the Change with the reas why it was unable to be completed and what the next steps for the requestor are.
2. Fill in the prompts as instructed. Change #, Change Status, Reason, Summary and the Requestor's email.
3. Confirm email has been sent.