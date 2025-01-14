# Veeam Report Email

Script to send Veeam Backup & Replication status report by email.

## How to Configure

1. Create a directory to store the script and reports. (You can configure the script to save the reports in this directory.)
2. Create a routine in Windows Task Schedule to run the script at the time you want to receive the report.
    2.1. Open Task Schedule - Action - Create Task
    2.2.  In Trigger, set the time the script should be executed.
    2.3.  In Action, Select the action "Start a program" and select the script file
