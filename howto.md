## Save Windows Server Print Logs in MySQL database.

### 1st part
You must enable the print logs and also the option to display the name of the printed file in the log file.

Open the event viewer and find the option:
Application and Services "_Log->Microsoft->Windows->PrintService->Operational_"

Click Properties->Check the "Enable Logs option and set the size of the log according to its structure. Leave the "Overwrite events as needed (oldest event first)

Access the "gpedit.msc" on the server where they are installed as printers and find an option: "_Computer Configuration->Administrative Templates->Printers->"Allow job name in event logs"_

Access the print manager, on printers, click on properties and on the "sharing" tab disable the "Process print jobs on client computers"

On some versions of Windows Server it may be necessary to enable the option through the registry: [HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Policies\Microsoft\Windows NT\Printers]
"ShowJobTitleInEventLogs" = dword: 00000001
If the document name does not appear in the logs, you may need to install an update: https://support.microsoft.com/kb/2938013

After performing the above steps you need to do some prints and check the log file to see if the file name is being displayed as expected.


### 2nd part
You must attach a task to the events log and pass this information to the powershell script, which will grab the print data and save it to the MySQL database.

Returning to the event viewer, on the same screen where the logs were activated, click on the EventID 307 and click on the option "Attach Task to this Event" and configure.

 - Define a name to the task. Example: "Print Log 307"
 - Action:  Start a program
 - Start a program: powershell.exe
 - Finish
 
Now you need to access the Task Manager and browse to the previously registered task, right-click on the task and click "Export" and save to any folder.

For simplicity, the XML file "EventID307_PrintLog.xml" is already changed with the correct settings, so you can import this file directly.

You must delete the job that was exported, right-click and import the XML file "EventID307_PrintLog.xml".

If you imported the task again, you need to adjust some parameters by selecting the task and right-clicking properties.
- In the General tab, check the option: "Run while the user is connected or not" and also "Run with higher privileges"
- On the Conditions tab, uncheck the option: "Start the task only if the computer is connected to the mains"
- On the Settings tab, change the option: "If the task is already running, the following rule will be applied" to "Put a new instance in the queue"

### 3rd part
If everything runs OK the script will generate a text file with printed jobs and also will save data in the MySQL database. Now that the data is in the database, just carry out the queries according to the need.

### 4th part
Usually the impressions made by "Microsoft Word" when using the option to print multiple copies, they are not counted correctly, so you need to run the "microsoft_office_word_forcesetcopycount.ps1" powershell script on the workstations to force the number of pages to be generated correctly .


### Sources:
- http://www.analistadeti.com/print-server-gerar-evento-de-impressao-event-viewer/
- https://gallery.technet.microsoft.com/Script-to-generate-print-84bdcf69
- http://www.thomasmaurer.ch/2011/04/powershell-run-mysql-querys-with-powershell/
- https://blogs.technet.microsoft.com/wincat/2011/08/25/trigger-a-powershell-script-from-a-windows-event/
- https://support.microsoft.com/en-us/kb/919736
