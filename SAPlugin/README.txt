MailEnable SpamAssassin(tm) Plugin
Date: 1st June 2003

What is it?
~~~~~~~~~~~
This plugin is a simple application that will process all messages going
through MailEnable via SpamAssassin software. It is not the most optimal 
way of doing it, since it just uses the pickup event of the MTA (MailEnable 
has a more advanced filtering plugin method, but time prevents me from doing
this). Proper testing has not been done, so try it out first before 
using it.

The source code is included in the archive, so you can modify, suggest
changes, fix bugs, etc. It can also be used as an example to write your
own pickup events. The sourcecode requires VB6.

For details about SpamAssassin, visit http://www.spamassassin.org.

MailEnable cannot provide any support for this plugin or SpamAssassin. 
While I am an employee, it was done in my own time. The MailEnable 
forum at http://forum.mailenable.com would be the best place to ask any 
questions. I'd regard this plugin as purely beta.


Performance
~~~~~~~~~~~
On a dual 733Mhz PIII, it can process about 1 spam a second (86400 messages
a day) and uses 100% CPU for the whole time.

There are lots of ways of improving performance. You might want to check
out whether you can get spamc/spamd going. And also implement whitelisting
and other features specific to the SpamAssassin software.


How to install
~~~~~~~~~~~~~~
First you need to install SpamAssassin. I used a Windows port of this:

http://sa.reusch.net

Just download the ZIP file and uncompress to a directory on your hard drive.

You can also do it all yourself by having a look at:

http://www.openhandhome.com/howtosa.html

Take the SAPlugin.exe and spamassassignconfig.ini files and copy to the 
Mail Enable\bin directory. Edit the .ini file to match your system, 
according to the Configuration details in this document. You now need to 
add this to the MTA pickup event. Open the MailEnable administration 
program, expand the Servers->Localhost->Agents branch, right click on the
MTA icon, and select Properties from the popup menu. Enable the pickup 
event and select the SAPlugin.exe file.

The MTA Pickup event is multithreaded. The SpamAssassin software doesn't
seem to like this too much (I don't know whether it is the PerlApp 
conversion, temp files the Perl app uses, something in the plugin, something 
in the MTA, etc.), so it is a lot faster and more reliable to run the MTA as
single-threaded. To do this, you need to change the following registry
setting:

HKEY_LOCAL_MACHINE\Software\Mail Enable\Mail Enable\Agents\MTA

The item "Maximum Transfer Threads" MUST be set to 1

You will need to restart the MTA service to enable this change. 


Configuration
~~~~~~~~~~~~~
The config file needs to reside in the same directory as the SAPlugin.exe
file. It is called spamassassinconfig.ini and an example of the contents
are below:

[SPAMASSASSIN Plugin Config]
QUARANTINE=1
QUARANTINEPATH=c:\progra~1\mailen~1\Quarantine
KEEPORIGINAL=0
SPAMASSASSINEXE=c:\spamas~1\spamassassin.exe
SPAMASSASSINRULESPATH=c:\spamas~1\rules
TEMPPATH=c:\progra~1\mailen~1\SATemp
MAXMESSAGESIZE=100000

QUARANTINE=1

0=Tag the email as spam
1=Quarantine spam detected
2=Delete the email (WARNING! This can/may/will delete legitimate email)

QUARANTINEPATH=c:\progra~1\mailen~1\Quarantine

Path to the Quarantine directory where the original message, the SpamAssassin
altered message, and the MailEnable command file is kept. The SpamAssassin
email has the extension of ".sa".

KEEPORIGINAL=0

0=The SpamAssassin modified email is passed to user, when no spam detected
1=When no spam detected, the original message is used

SPAMASSASSINEXE=c:\spamas~1\spamassassin.exe

Full path to the SpamAssassin executable.

SPAMASSASSINRULESPATH=c:\spamas~1\rules

Full path to the SpamAssassign rules directory.

TEMPPATH=c:\progra~1\mailen~1\SATemp

Path to the temp directory that will be used for processing.

MAXMESSAGESIZE=100000

Any message over this limit is not checked for spam. Since most spams
are relatively small, this will speed up processing. The size is in
bytes. Use 0 in order to check every email.


---
SpamAssassin is a trademark of Deersoft, Inc.


