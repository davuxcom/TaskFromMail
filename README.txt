Initial: 8/25/2011
Update:  9/17/2011
DSA v1.0.4

TaskFromMail Description:

This Outlook Add-In adds two buttons to the right-click context 
menu in Outlook ("Create Task..." and "Append to task...").
Clicking the Create Task button will create a new task, and attach  
the selected email message to the new task.  The subject line of 
the new task will reflect the subject line of the selected message. 
Clicking the "Append to task..." button will append the message to 
an existing task in the folder of your choosing.

Use Notes:

- Hold CTRL while clicking 'Create Task...' or 'Append to task...'
  to bypass default and force the task folder selection window.

- Alternatively, ClearSettings.reg will remove any saved settings.

- On the 'Select Task' dialog, items are grouped according to their
  first category only.
--------------------------------------------------------------

Install Notes:

- Run \Install\setup.exe

Upgrade Notes:

In order to upgrade, the install path must be exactly the same.

If you installed from C:\Users\Chris\Desktop\TFM1.0\Install 
then you MUST always use that install path.

Otherwise, the add-in will need to be uninstalled before 
reinstalling.  Settings will not be lost.

--------------------------------------------------------------

Requirements:
- Office 2007 or Office 2010 (x86 or x64)
- Windows XP, Vista or 7 (x86 or x64)

Prerequisites:
- .NET 3.5sp1 Full Profile
- Visual Studio Tools for Office (VSTO)
- Windows Installer 3.1

The prerequisites will be downloaded from Microsoft and installed 
by setup.exe if the user has permission to do so.

--------------------------------------------------------------

Tested against:

Windows 7 x64, Office 2010 x86
Windows XP x86, Office 2007 x86

Settings location:

[HKEY_CURRENT_USER\Software\DSASoftware\TaskFromMail]
"DefaultGroup"=""
"DefaultFolder"=""

Troubleshooting instructions:

- Open a command prompt (start -> run -> cmd)
- Enter "set VSTO_SUPPRESSDISPLAYALERTS=0" 
  without quotes, and press enter.
- Enter "c:\Program Files\Microsoft Office\Office12\OUTLOOK.EXE"
  _with_ quotes, and press enter.
- Outlook will open in VSTO debug mode.
- A box will open indicating there is a problem, click 'Details'
  copy and paste the text block and send to the support staff.

Changelog:

v1.0.0
	Initial build
v1.0.3
	Added task folder selection Window
v1.0.4
	Added 'Append to task...' option
