Outlook_Connection_Checker
--------------------------
v1.0.0 - Initial Release

v1.0.1 - Fixes to check methods to make it work in cached and none cached mode.

v1.0.2 - Added 'Clogged outbox" check

The Problem:
------------
Outlook is not good at reporting to the user if they are working offline due to either offline mode, or connectivity issues.
Outlook only warns the user by changing its status bar and placing a warning sign on the icon, which can go ignored.
I have seen many occasions where a user has built up emails in their outbox throughout the day and not even noticed they havent been getting messages, and theirs haven't been being delivered.
They carry on working oblivious to their lack of productivity.

The Solution:
-------------
A Simple .net 4.8 executbale which upon running checks if Outlook is running for the user, if not it closes and does nothing.
If so it checks if Outlook is working in offline mode, or has connectivity problems to M365 (or your own EWS service).
It gives the user a meaningfull messagebox as well as a toast notification then exits.
The user can remediate, or knows to contact their helpdesk.  
Another feature I have added since is that from v1.0.2 it will count the messages in the outbox.  If ther eare 5 messages or more it will wait 4 minutes then count them again.  If stil 5 or more it will warn the user.

Usage
-----
Deploy .net framework 4.8 and depoly the windwos binaries to a folder on the users machine or the RDS farm.
Create a scheduled task running Outlook_Connection_Checker.exe as the user from Login of the user repeating every X minutes (5, 15, 60, whatever works best for you, I use 15, dotn recommend less than 5) for the lengh of a working day or more.
All of this can easily be automated via Group Policy, or Intune/Endpoint Manager or be done manually if you just want to setup for 1 or 2 users.

Have peace of mind that if your users are not connected to their email service they will know about it in a reasonable time.  

To Do
-----

Create a config file so that features can be enabled/disabled, and time between outbox check can be configured. 
Encorporate the age of the messages in the outbox to the cehck, rather than the number.  
