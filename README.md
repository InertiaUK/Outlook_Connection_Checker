Outlook_Connection_Checker
--------------------------
v1.0.0

The Problen:
------------
Outlook is not good at reporting to the user if they are working offline due to either offline mode, or connectivity issues.
Outlook only warns the user by changing its status bar and placing a warning sign on the icon, which can go ignored.
I have seen many occasions where a user has built up emails in their outbox throughout the day and not even noticed they havent been getting messages, and theirs haven't been being delivered.
They carry on working oblivious to their lack of productivity.

The Solution:
-------------
A Simple .net 4.8 executbale which upon running checks if Outlook is running for the user, if not it closes and does nothing.
If so it checks if Outlook is working in offline mode, or has connectivity problems to M365 (or your own EWS service).
It gives the user a meaningful messagebox as well as a toast notification then exits.
The user can remediate, or knows to contact their helpdesk.  

Usage
-----
Deploy .net framework 4.8 and depoly the windwos binaries to a folder on the users machine or the RDS farm.
Create a scheduled task running Outlook_Connection_Checker.exe as the user from Login of the user repeating every X minutes (5, 15, 60, whatever works best for you, I use 15) for the lengh of a working day or more.
All of this can easily be automated via Group Policy, or Intune/Endpoint Manager or be doen manually if you just want to setup for 1 or 2 users.

Have peace of mind that if your users are not connected to their email service they will know about it in a reasonable time.  
