#Requires -Version 3.0
#Requires -PSEdition Desktop
#Requires -Module @{ModuleName="Veeam.Backup.PowerShell"; ModuleVersion="1.0"}

#region User-Variables
# VBR Server (Server Name)
$vbrServer = (Get-WmiObject win32_computersystem).DNSHostName+"."+(Get-WmiObject win32_computersystem).Domain
# Report mode (RPO) - valid modes: any number of hours, Weekly or Monthly
# 24, 48, "Weekly", "Monthly"
$reportMode = 24
# Report Title - Name that will be displayed in the report header
$rptTitle = "<Report Title - Change this>"
# HTML Report Width (Percent)
$rptWidth = 97


# HTML File output path and filename
$pathHTML = "V:\Report\VeeamReport_$(Get-Date -format dd-MM-yyyy).htm"

# wkhtmlpdf path (Used to convert html to PDF)
$wkhtmltopdfPath = "C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"

# PDF File output path and filename
$pathPDF = "V:\Report\Report Last 24h of backup_$(Get-Date -format dd-MM-yyyy).pdf"

# Show VM Backup Protection Summary (across entire infrastructure)
$showSummaryProtect = $true
# Show VMs with No Successful Backups within RPO ($reportMode)
$showUnprotectedVMs = $true
# Show VMs with Successful Backups within RPO ($reportMode)
# Also shows VMs with Only Backups with Warnings within RPO ($reportMode)
$showProtectedVMs = $true
# Exclude VMs from Missing and Successful Backups sections
# $excludevms = @("")
$excludeVMs = @("")
# Exclude VMs from Missing and Successful Backups sections in the following (vCenter) folder(s)
# $excludeFolder = @("folder1","folder2","*_testonly")
$excludeFolder = @("")
# Exclude VMs from Missing and Successful Backups sections in the following (vCenter) datacenter(s)
# $excludeDC = @("dc1","dc2","dc*")
$excludeDC = @("")
# Exclude VMs from Missing and Successful Backups sections in the following (vCenter) clusters
$excludeCluster = @("")
# Exclude Templates from Missing and Successful Backups sections
$excludeTemp = $true

# Show VMs Backed Up by Multiple Jobs within time frame ($reportMode)
$showMultiJobs = $true

# Show Backup Session Summary
$showSummaryBk = $true
# Show Backup Job Status
$showJobsBk = $true
# Show Backup Job Size (total)
$showBackupSizeBk = $false
# Show detailed information for Backup Jobs/Sessions (Avg Speed, Total(GB), Processed(GB), Read(GB), Transferred(GB), Dedupe, Compression)
$showDetailedBk = $true
# Show all Backup Sessions within time frame ($reportMode)
$showAllSessBk = $true
# Show all Backup Tasks from Sessions within time frame ($reportMode)
$showAllTasksBk = $false
# Show Running Backup Jobs
$showRunningBk = $true
# Show Running Backup Tasks
$showRunningTasksBk = $false
# Show Backup Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFailBk = $true
# Show Backup Tasks w/Warnings or Failures from Sessions within time frame ($reportMode)
$showTaskWFBk = $false
# Show Successful Backup Sessions within time frame ($reportMode)
$showSuccessBk = $false
# Show Successful Backup Tasks from Sessions within time frame ($reportMode)
$showTaskSuccessBk = $false
# Only show last Session for each Backup Job
$onlyLastBk = $true
# Only report on the following Backup Job(s)
#$backupJob = @("Backup Job 1","Backup Job 3","Backup Job *")
$backupJob = @("")

# Show Running Restore VM Sessions
$showRestoRunVM = $true
# Show Completed Restore VM Sessions within time frame ($reportMode)
$showRestoreVM = $true

# Show Replication Session Summary
$showSummaryRp = $false
# Show Replication Job Status
$showJobsRp = $true
# Show detailed information for Replication Jobs/Sessions (Avg Speed, Total(GB), Processed(GB), Read(GB), Transferred(GB), Dedupe, Compression)
$showDetailedRp = $true
# Show all Replication Sessions within time frame ($reportMode)
$showAllSessRp = $false
# Show all Replication Tasks from Sessions within time frame ($reportMode)
$showAllTasksRp = $false
# Show Running Replication Jobs
$showRunningRp = $false
# Show Running Replication Tasks
$showRunningTasksRp = $false
# Show Replication Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFailRp = $false
# Show Replication Tasks w/Warnings or Failures from Sessions within time frame ($reportMode)
$showTaskWFRp = $false
# Show Successful Replication Sessions within time frame ($reportMode)
$showSuccessRp = $false
# Show Successful Replication Tasks from Sessions within time frame ($reportMode)
$showTaskSuccessRp = $false
# Only show last session for each Replication Job
$onlyLastRp = $false
# Only report on the following Replication Job(s)
#$replicaJob = @("Replica Job 1","Replica Job 3","Replica Job *")
$replicaJob = @("")

# Show Backup Copy Session Summary
$showSummaryBc = $true
# Show Backup Copy Job Status
$showJobsBc = $true
# Show Backup Copy Job Size (total)
$showBackupSizeBc = $false
# Show detailed information for Backup Copy Sessions (Avg Speed, Total(GB), Processed(GB), Read(GB), Transferred(GB), Dedupe, Compression)
$showDetailedBc = $true
# Show all Backup Copy Sessions within time frame ($reportMode)
$showAllSessBc = $false
# Show all Backup Copy Tasks from Sessions within time frame ($reportMode)
$showAllTasksBc = $false
# Show Idle Backup Copy Sessions
$showIdleBc = $false
# Show Pending Backup Copy Tasks
$showPendingTasksBc = $false
# Show Working Backup Copy Jobs
$showRunningBc = $false
# Show Working Backup Copy Tasks
$showRunningTasksBc = $false
# Show Backup Copy Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFailBc = $false
# Show Backup Copy Tasks w/Warnings or Failures from Sessions within time frame ($reportMode)
$showTaskWFBc = $false
# Show Successful Backup Copy Sessions within time frame ($reportMode)
$showSuccessBc = $false
# Show Successful Backup Copy Tasks from Sessions within time frame ($reportMode)
$showTaskSuccessBc = $false
# Only show last Session for each Backup Copy Job
$onlyLastBc = $false
# Only report on the following Backup Copy Job(s)
#$bcopyJob = @("Backup Copy Job 1","Backup Copy Job 3","Backup Copy Job *")
$bcopyJob = @("")

# Show Tape Backup Session Summary
$showSummaryTp = $false
# Show Tape Backup Job Status
$showJobsTp = $true
# Show detailed information for Tape Backup Sessions (Avg Speed, Total(GB), Read(GB), Transferred(GB))
$showDetailedTp = $true
# Show all Tape Backup Sessions within time frame ($reportMode)
$showAllSessTp = $true
# Show all Tape Backup Tasks from Sessions within time frame ($reportMode)
$showAllTasksTp = $true
# Show Waiting Tape Backup Sessions
$showWaitingTp = $false
# Show Idle Tape Backup Sessions
$showIdleTp = $false
# Show Pending Tape Backup Tasks
$showPendingTasksTp = $false
# Show Working Tape Backup Jobs
$showRunningTp = $false
# Show Working Tape Backup Tasks
$showRunningTasksTp = $false
# Show Tape Backup Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFailTp = $false
# Show Tape Backup Tasks w/Warnings or Failures from Sessions within time frame ($reportMode)
$showTaskWFTp = $false
# Show Successful Tape Backup Sessions within time frame ($reportMode)
$showSuccessTp = $false
# Show Successful Tape Backup Tasks from Sessions within time frame ($reportMode)
$showTaskSuccessTp = $false
# Only show last Session for each Tape Backup Job
$onlyLastTp = $false
# Only report on the following Tape Backup Job(s)
#$tapeJob = @("Tape Backup Job 1","Tape Backup Job 3","Tape Backup Job *")
$tapeJob = @("")

# Show all Tapes
$showTapes = $false
# Show all Tapes by (Custom) Media Pool
$showTpMp = $false
# Show all Tapes by Vault
$showTpVlt = $false
# Show all Expired Tapes
$showExpTp = $false
# Show Expired Tapes by (Custom) Media Pool
$showExpTpMp = $false
# Show Expired Tapes by Vault
$showExpTpVlt = $false
# Show Tapes written to within time frame ($reportMode)
$showTpWrt = $false

# Show Agent Backup Session Summary
$showSummaryEp = $false
# Show Agent Backup Job Status
$showJobsEp = $false
# Show Agent Backup Job Size (total)
$showBackupSizeEp = $false
# Show all Agent Backup Sessions within time frame ($reportMode)
$showAllSessEp = $false
# Show Running Agent Backup jobs
$showRunningEp = $false
# Show Agent Backup Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFailEp = $false
# Show Successful Agent Backup Sessions within time frame ($reportMode)
$showSuccessEp = $false
# Only show last session for each Agent Backup Job
$onlyLastEp = $false
# Only report on the following Agent Backup Job(s)
#$epbJob = @("Agent Backup Job 1","Agent Backup Job 3","Agent Backup Job *")
$epbJob = @("")

# Show SureBackup Session Summary
$showSummarySb = $false
# Show SureBackup Job Status
$showJobsSb = $false
# Show all SureBackup Sessions within time frame ($reportMode)
$showAllSessSb = $false
# Show all SureBackup Tasks from Sessions within time frame ($reportMode)
$showAllTasksSb = $false
# Show Running SureBackup Jobs
$showRunningSb = $false
# Show Running SureBackup Tasks
$showRunningTasksSb = $false
# Show SureBackup Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFailSb = $false
# Show SureBackup Tasks w/Warnings or Failures from Sessions within time frame ($reportMode)
$showTaskWFSb = $false
# Show Successful SureBackup Sessions within time frame ($reportMode)
$showSuccessSb = $false
# Show Successful SureBackup Tasks from Sessions within time frame ($reportMode)
$showTaskSuccessSb = $false
# Only show last Session for each SureBackup Job
$onlyLastSb = $false
# Only report on the following SureBackup Job(s)
#$surebJob = @("SureBackup Job 1","SureBackup Job 3","SureBackup Job *")
$surebJob = @("")

# Show Configuration Backup Summary
$showSummaryConfig = $true
# Show Proxy Info
$showProxy = $true
# Show Repository Info
$showRepo = $true
# Show Repository Permissions for Agent Jobs
$showRepoPerms = $false
# Show Replica Target Info
$showReplicaTarget = $false
# Show Veeam Services Info (Windows Services)
$showServices = $true
# Show only Services that are NOT running
$hideRunningSvc = $true
# Show License expiry info
$showLicExp = $true

# Highlighting Thresholds
# Repository Free Space Remaining %
$repoCritical = 10
$repoWarn = 20
# Replica Target Free Space Remaining %
$replicaCritical = 10
$replicaWarn = 20
# License Days Remaining
$licenseCritical = 30
$licenseWarn = 90
#endregion

#Veeam Server Version
# $VeeamVersion

$CoreDllPath = (Get-ItemProperty -Path "HKLM:\Software\Veeam\Veeam Backup and Replication\" | Select-Object -ExpandProperty CorePath) + "Veeam.Backup.Core.dll"
$CoreDll = Get-Item -Path $CoreDllPath
$VeeamVersion = $CoreDll.VersionInfo.ProductVersion

#region Connect
Import-Module Veeam.Backup.PowerShell -WarningAction SilentlyContinue

# Connect to VBR server
$OpenConnection = (Get-VBRServerSession).Server
If ($OpenConnection -ne $vbrServer) {
    Write-Verbose "Connecting to $vbrServer."
    Disconnect-VBRServer
    Connect-VBRServer -Server $vbrServer -ErrorAction Stop
} else {
    Write-Verbose "Already connected to $vbrServer."
}
#endregion

#region NonUser-Variables
# Get all Backup/Backup Copy/Replica Jobs
$allJobs = @()
If ($showSummaryBk + $showJobsBk + $showAllSessBk + $showAllTasksBk + $showRunningBk +
    $showRunningTasksBk + $showWarnFailBk + $showTaskWFBk + $showSuccessBk + $showTaskSuccessBk +
    $showSummaryRp + $showJobsRp + $showAllSessRp + $showAllTasksRp + $showRunningRp +
    $showRunningTasksRp + $showWarnFailRp + $showTaskWFRp + $showSuccessRp + $showTaskSuccessRp +
    $showSummaryBc + $showJobsBc + $showAllSessBc + $showAllTasksBc + $showIdleBc +
    $showPendingTasksBc + $showRunningBc + $showRunningTasksBc + $showWarnFailBc +
    $showTaskWFBc + $showSuccessBc + $showTaskSuccessBc) {
    $allJobs = Get-VBRJob -WarningAction SilentlyContinue
}
# Get all Backup Jobs
$allJobsBk = @($allJobs | Where-Object { $_.JobType -eq "Backup" })
# Get all Replication Jobs
$allJobsRp = @($allJobs | Where-Object { $_.JobType -eq "Replica" })
# Get all Backup Copy Jobs
$allJobsBc = @($allJobs | Where-Object { $_.JobType -eq "BackupSync" })
# Get all Tape Jobs
$allJobsTp = @()
If ($showSummaryTp + $showJobsTp + $showAllSessTp + $showAllTasksTp +
    $showWaitingTp + $showIdleTp + $showPendingTasksTp + $showRunningTp + $showRunningTasksTp +
    $showWarnFailTp + $showTaskWFTp + $showSuccessTp + $showTaskSuccessTp) {
    $allJobsTp = @(Get-VBRTapeJob)
}
# Get all Agent Backup Jobs
$allJobsEp = @()
If ($showSummaryEp + $showJobsEp + $showAllSessEp + $showRunningEp +
    $showWarnFailEp + $showSuccessEp) {
    $allJobsEp = @(Get-VBREPJob)
}
# Get all SureBackup Jobs
$allJobsSb = @()
If ($showSummarySb + $showJobsSb + $showAllSessSb + $showAllTasksSb +
    $showRunningSb + $showRunningTasksSb + $showWarnFailSb + $showTaskWFSb +
    $showSuccessSb + $showTaskSuccessSb) {
    $allJobsSb = @(Get-VSBJob)
}

# Get all Backup/Backup Copy/Replica Sessions
$allSess = @()
If ($allJobs) {
    $allSess = Get-VBRBackupSession
}
# Get all Restore Sessions
$allSessResto = @()
If ($showRestoRunVM + $showRestoreVM) {
    $allSessResto = Get-VBRRestoreSession
}
# Get all Tape Backup Sessions
$allSessTp = @()
If ($allJobsTp) {
    Foreach ($tpJob in $allJobsTp) {
        $tpSessions = [veeam.backup.core.cbackupsession]::GetByJob($tpJob.id)
        $allSessTp += $tpSessions
    }
}
# Get all Agent Backup Sessions
$allSessEp = @()
If ($allJobsEp) {
    $allSessEp = Get-VBREPSession
}
# Get all SureBackup Sessions
$allSessSb = @()
If ($allJobsSb) {
    $allSessSb = Get-VSBSession
}

# Get all Backups
$jobBackups = @()
If ($showBackupSizeBk + $showBackupSizeBc + $showBackupSizeEp) {
    $jobBackups = Get-VBRBackup
}
# Get Backup Job Backups
$backupsBk = @($jobBackups | Where-Object { $_.JobType -eq "Backup" })
# Get Backup Copy Job Backups
$backupsBc = @($jobBackups | Where-Object { $_.JobType -eq "BackupSync" })
# Get Agent Backup Job Backups
$backupsEp = @($jobBackups | Where-Object { $_.JobType -eq "EndpointBackup" })

# Get all Media Pools
$mediaPools = Get-VBRTapeMediaPool
# Get all Media Vaults
$mediaVaults = Get-VBRTapeVault
# Get all Tapes
$mediaTapes = Get-VBRTapeMedium
# Get all Tape Libraries
$mediaLibs = Get-VBRTapeLibrary
# Get all Tape Drives
$mediaDrives = Get-VBRTapeDrive

# Get Configuration Backup Info
$configBackup = Get-VBRConfigurationBackupJob
# Get VBR Server object
$vbrServerObj = Get-VBRLocalhost
# Get all Proxies
$proxyList = Get-VBRViProxy
# Get all Repositories
$repoList = Get-VBRBackupRepository
$repoListSo = Get-VBRBackupRepository -ScaleOut
# Get all Tape Servers
$tapesrvList = Get-VBRTapeServer

# Convert mode (timeframe) to hours
If ($reportMode -eq "Monthly") {
    $HourstoCheck = 720
} Elseif ($reportMode -eq "Weekly") {
    $HourstoCheck = 168
} Else {
    $HourstoCheck = $reportMode
}

# Gather all Backup Sessions within timeframe
$sessListBk = @($allSess | Where-Object { ($_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -eq "Working") -and $_.JobType -eq "Backup" })
If ($null -ne $backupJob -and $backupJob -ne "") {
    $allJobsBkTmp = @()
    $sessListBkTmp = @()
    $backupsBkTmp = @()
    Foreach ($bkJob in $backupJob) {
        $allJobsBkTmp += $allJobsBk | Where-Object { $_.Name -like $bkJob }
        $sessListBkTmp += $sessListBk | Where-Object { $_.JobName -like $bkJob }
        $backupsBkTmp += $backupsBk | Where-Object { $_.JobName -like $bkJob }
    }
    $allJobsBk = $allJobsBkTmp | Sort-Object Id -Unique
    $sessListBk = $sessListBkTmp | Sort-Object Id -Unique
    $backupsBk = $backupsBkTmp | Sort-Object Id -Unique
}
If ($onlyLastBk) {
    $tempSessListBk = $sessListBk
    $sessListBk = @()
    Foreach ($job in $allJobsBk) {
        $sessListBk += $tempSessListBk | Where-Object { $_.Jobname -eq $job.name } | Sort-Object EndTime -Descending | Select-Object -First 1
    }
}
# Get Backup Session information
$totalXferBk = 0
$totalReadBk = 0
$sessListBk | ForEach-Object { $totalXferBk += $([Math]::Round([Decimal]$_.Progress.TransferedSize / 1GB, 2)) }
$sessListBk | ForEach-Object { $totalReadBk += $([Math]::Round([Decimal]$_.Progress.ReadSize / 1GB, 2)) }
$successSessionsBk = @($sessListBk | Where-Object { $_.Result -eq "Success" })
$warningSessionsBk = @($sessListBk | Where-Object { $_.Result -eq "Warning" })
$failsSessionsBk = @($sessListBk | Where-Object { $_.Result -eq "Failed" })
$runningSessionsBk = @($sessListBk | Where-Object { $_.State -eq "Working" })
$failedSessionsBk = @($sessListBk | Where-Object { ($_.Result -eq "Failed") -and ($_.WillBeRetried -ne "True") })

# Gather VM Restore Sessions within timeframe
$sessListResto = @($allSessResto | Where-Object { $_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or !($_.IsCompleted) })
# Get VM Restore Session information
$completeResto = @($sessListResto | Where-Object { $_.IsCompleted })
$runningResto = @($sessListResto | Where-Object { !($_.IsCompleted) })

# Gather all Replication Sessions within timeframe
$sessListRp = @($allSess | Where-Object { ($_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -eq "Working") -and $_.JobType -eq "Replica" })
If ($null -ne $replicaJob -and $replicaJob -ne "") {
    $allJobsRpTmp = @()
    $sessListRpTmp = @()
    Foreach ($rpJob in $replicaJob) {
        $allJobsRpTmp += $allJobsRp | Where-Object { $_.Name -like $rpJob }
        $sessListRpTmp += $sessListRp | Where-Object { $_.JobName -like $rpJob }
    }
    $allJobsRp = $allJobsRpTmp | Sort-Object Id -Unique
    $sessListRp = $sessListRpTmp | Sort-Object Id -Unique
}
If ($onlyLastRp) {
    $tempSessListRp = $sessListRp
    $sessListRp = @()
    Foreach ($job in $allJobsRp) {
        $sessListRp += $tempSessListRp | Where-Object { $_.Jobname -eq $job.name } | Sort-Object EndTime -Descending | Select-Object -First 1
    }
}
# Get Replication Session information
$totalXferRp = 0
$totalReadRp = 0
$sessListRp | ForEach-Object { $totalXferRp += $([Math]::Round([Decimal]$_.Progress.TransferedSize / 1GB, 2)) }
$sessListRp | ForEach-Object { $totalReadRp += $([Math]::Round([Decimal]$_.Progress.ReadSize / 1GB, 2)) }
$successSessionsRp = @($sessListRp | Where-Object { $_.Result -eq "Success" })
$warningSessionsRp = @($sessListRp | Where-Object { $_.Result -eq "Warning" })
$failsSessionsRp = @($sessListRp | Where-Object { $_.Result -eq "Failed" })
$runningSessionsRp = @($sessListRp | Where-Object { $_.State -eq "Working" })
$failedSessionsRp = @($sessListRp | Where-Object { ($_.Result -eq "Failed") -and ($_.WillBeRetried -ne "True") })

# Gather all Backup Copy Sessions within timeframe
$sessListBc = @($allSess | Where-Object { ($_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -match "Working|Idle") -and $_.JobType -eq "BackupSync" })
If ($null -ne $bcopyJob -and $bcopyJob -ne "") {
    $allJobsBcTmp = @()
    $sessListBcTmp = @()
    $backupsBcTmp = @()
    Foreach ($bcJob in $bcopyJob) {
        $allJobsBcTmp += $allJobsBc | Where-Object { $_.Name -like $bcJob }
        $sessListBcTmp += $sessListBc | Where-Object { $_.JobName -like $bcJob }
        $backupsBcTmp += $backupsBc | Where-Object { $_.JobName -like $bcJob }
    }
    $allJobsBc = $allJobsBcTmp | Sort-Object Id -Unique
    $sessListBc = $sessListBcTmp | Sort-Object Id -Unique
    $backupsBc = $backupsBcTmp | Sort-Object Id -Unique
}
If ($onlyLastBc) {
    $tempSessListBc = $sessListBc
    $sessListBc = @()
    Foreach ($job in $allJobsBc) {
        $sessListBc += $tempSessListBc | Where-Object { $_.Jobname -eq $job.name -and $_.BaseProgress -eq 100 } | Sort-Object EndTime -Descending | Select-Object -First 1
    }
}
# Get Backup Copy Session information
$totalXferBc = 0
$totalReadBc = 0
$sessListBc | ForEach-Object { $totalXferBc += $([Math]::Round([Decimal]$_.Progress.TransferedSize / 1GB, 2)) }
$sessListBc | ForEach-Object { $totalReadBc += $([Math]::Round([Decimal]$_.Progress.ReadSize / 1GB, 2)) }
$idleSessionsBc = @($sessListBc | Where-Object { $_.State -eq "Idle" })
$successSessionsBc = @($sessListBc | Where-Object { $_.Result -eq "Success" })
$warningSessionsBc = @($sessListBc | Where-Object { $_.Result -eq "Warning" })
$failsSessionsBc = @($sessListBc | Where-Object { $_.Result -eq "Failed" })
$workingSessionsBc = @($sessListBc | Where-Object { $_.State -eq "Working" })

# Gather all Tape Backup Sessions within timeframe
$sessListTp = @($allSessTp | Where-Object { $_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -match "Working|Idle" })
If ($null -ne $tapeJob -and $tapeJob -ne "") {
    $allJobsTpTmp = @()
    $sessListTpTmp = @()
    Foreach ($tpJob in $tapeJob) {
        $allJobsTpTmp += $allJobsTp | Where-Object { $_.Name -like $tpJob }
        $sessListTpTmp += $sessListTp | Where-Object { $_.JobName -like $tpJob }
    }
    $allJobsTp = $allJobsTpTmp | Sort-Object Id -Unique
    $sessListTp = $sessListTpTmp | Sort-Object Id -Unique
}
If ($onlyLastTp) {
    $tempSessListTp = $sessListTp
    $sessListTp = @()
    Foreach ($job in $allJobsTp) {
        $sessListTp += $tempSessListTp | Where-Object { $_.Jobname -eq $job.name } | Sort-Object EndTime -Descending | Select-Object -First 1
    }
}
# Get Tape Backup Session information
$totalXferTp = 0
$totalReadTp = 0
$sessListTp | ForEach-Object { $totalXferTp += $([Math]::Round([Decimal]$_.Progress.TransferedSize / 1GB, 2)) }
$sessListTp | ForEach-Object { $totalReadTp += $([Math]::Round([Decimal]$_.Progress.ReadSize / 1GB, 2)) }
$idleSessionsTp = @($sessListTp | Where-Object { $_.State -eq "Idle" })
$successSessionsTp = @($sessListTp | Where-Object { $_.Result -eq "Success" })
$warningSessionsTp = @($sessListTp | Where-Object { $_.Result -eq "Warning" })
$failsSessionsTp = @($sessListTp | Where-Object { $_.Result -eq "Failed" })
$workingSessionsTp = @($sessListTp | Where-Object { $_.State -eq "Working" })
$waitingSessionsTp = @($sessListTp | Where-Object { $_.State -eq "WaitingTape" })

# Gather all Agent Backup Sessions within timeframe
$sessListEp = $allSessEp | Where-Object { ($_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -eq "Working") }
If ($null -ne $epbJob -and $epbJob -ne "") {
    $allJobsEpTmp = @()
    $sessListEpTmp = @()
    $backupsEpTmp = @()
    Foreach ($eJob in $epbJob) {
        $allJobsEpTmp += $allJobsEp | Where-Object { $_.Name -like $eJob }
        $backupsEpTmp += $backupsEp | Where-Object { $_.JobName -like $eJob }
    }
    Foreach ($job in $allJobsEpTmp) {
        $sessListEpTmp += $sessListEp | Where-Object { $_.JobId -eq $job.Id }
    }
    $allJobsEp = $allJobsEpTmp | Sort-Object Id -Unique
    $sessListEp = $sessListEpTmp | Sort-Object Id -Unique
    $backupsEp = $backupsEpTmp | Sort-Object Id -Unique
}
If ($onlyLastEp) {
    $tempSessListEp = $sessListEp
    $sessListEp = @()
    Foreach ($job in $allJobsEp) {
        $sessListEp += $tempSessListEp | Where-Object { $_.JobId -eq $job.Id } | Sort-Object EndTime -Descending | Select-Object -First 1
    }
}
# Get Agent Backup Session information
$successSessionsEp = @($sessListEp | Where-Object { $_.Result -eq "Success" })
$warningSessionsEp = @($sessListEp | Where-Object { $_.Result -eq "Warning" })
$failsSessionsEp = @($sessListEp | Where-Object { $_.Result -eq "Failed" })
$runningSessionsEp = @($sessListEp | Where-Object { $_.State -eq "Working" })

# Gather all SureBackup Sessions within timeframe
$sessListSb = @($allSessSb | Where-Object { $_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -ne "Stopped" })
If ($null -ne $surebJob -and $surebJob -ne "") {
    $allJobsSbTmp = @()
    $sessListSbTmp = @()
    Foreach ($SbJob in $surebJob) {
        $allJobsSbTmp += $allJobsSb | Where-Object { $_.Name -like $SbJob }
        $sessListSbTmp += $sessListSb | Where-Object { $_.JobName -like $SbJob }
    }
    $allJobsSb = $allJobsSbTmp | Sort-Object Id -Unique
    $sessListSb = $sessListSbTmp | Sort-Object Id -Unique
}
If ($onlyLastSb) {
    $tempSessListSb = $sessListSb
    $sessListSb = @()
    Foreach ($job in $allJobsSb) {
        $sessListSb += $tempSessListSb | Where-Object { $_.Jobname -eq $job.name } | Sort-Object EndTime -Descending | Select-Object -First 1
    }
}
# Get SureBackup Session information
$successSessionsSb = @($sessListSb | Where-Object { $_.Result -eq "Success" })
$warningSessionsSb = @($sessListSb | Where-Object { $_.Result -eq "Warning" })
$failsSessionsSb = @($sessListSb | Where-Object { $_.Result -eq "Failed" })
$runningSessionsSb = @($sessListSb | Where-Object { $_.State -ne "Stopped" })

# Format Report Mode for header
If (($reportMode -ne "Weekly") -And ($reportMode -ne "Monthly")) {
    $rptMode = "RPO: $reportMode Hrs"
} Else {
    $rptMode = "RPO: $reportMode"
}

# Append Report Mode to Email subject
If ($modeSubject) {
    If (($reportMode -ne "Weekly") -And ($reportMode -ne "Monthly")) {
        $emailSubject = "$emailSubject (Last $reportMode Hrs)"
    } Else {
        $emailSubject = "$emailSubject ($reportMode)"
    }
}

# Append Date and Time to Email subject
If ($dtSubject) {
    $emailSubject = "$emailSubject - $(Get-Date -format g)"
}
#endregion

#region Functions

Function Get-VBRProxyInfo {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, ValueFromPipeline = $true)]
        [PSObject[]]$Proxy
    )
    Begin {
        $outputAry = @()
        Function Initialize-Object {
            param (
                [PsObject]
                $InputObject
            )
            $ping = new-object system.net.networkinformation.ping
            $isIP = '\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b'
            If ($InputObject.Host.Name -match $isIP) {
                $IPv4 = $InputObject.Host.Name
            } Else {
                $DNS = [Net.DNS]::GetHostEntry("$($InputObject.Host.Name)")
                $IPv4 = ($DNS.get_AddressList() | Where-Object { $_.AddressFamily -eq "InterNetwork" } | Select-Object -First 1).IPAddressToString
            }
            $pinginfo = $ping.send("$($IPv4)")
            If ($pinginfo.Status -eq "Success") {
                $hostAlive = "Alive"
                $response = $pinginfo.RoundtripTime
            } Else {
                $hostAlive = "Dead"
                $response = $null
            }
            If ($InputObject.IsDisabled) {
                $enabled = "False"
            } Else {
                $enabled = "True"
            }
            $tMode = switch ($InputObject.Options.TransportMode) {
                "Auto" { "Automatic" }
                "San" { "Direct SAN" }
                "HotAdd" { "Hot Add" }
                "Nbd" { "Network" }
                default { "Unknown" }
            }
            $vPCFuncObject = New-Object PSObject -Property @{
                ProxyName = $InputObject.Name
                RealName  = $InputObject.Host.Name.ToLower()
                Disabled  = $InputObject.IsDisabled
                pType     = $InputObject.ChassisType
                Status    = $hostAlive
                IP        = $IPv4
                Response  = $response
                Enabled   = $enabled
                maxtasks  = $InputObject.Options.MaxTasksCount
                tMode     = $tMode
            }
            Return $vPCFuncObject
        }
    }
    Process {
        Foreach ($p in $Proxy) {
            $outputObj = Initialize-Object -InputObject $p
        }
        $outputAry += $outputObj
    }
    End {
        $outputAry
    }
}

Function Get-VBRRepoInfo {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, ValueFromPipeline = $true)]
        [PSObject[]]$Repository
    )
    Begin {
        $outputAry = @()
        Function Initialize-Object {
            param(
                $Name,
                $RepoHost,
                $Path,
                $Free,
                $Total,
                $MaxTasks,
                $RType
            )
            $repoObj = New-Object -TypeName PSObject -Property @{
                Target         = $name
                RepoHost       = $repohost
                Storepath      = $path
                StorageFree    = [Math]::Round([Decimal]$free / 1GB, 2)
                StorageTotal   = [Math]::Round([Decimal]$total / 1GB, 2)
                FreePercentage = [Math]::Round(($free / $total) * 100)
                MaxTasks       = $maxtasks
                rType          = $rtype
            }
            Return $repoObj
        }
    }
    Process {
        Foreach ($r in $Repository) {
            # Refresh Repository Size Info
            [Veeam.Backup.Core.CBackupRepositoryEx]::SyncSpaceInfoToDb($r, $true)
            $rType = switch ($r.Type) {
                "WinLocal" { "Windows Local" }
                "LinuxLocal" { "Linux Local" }
                "CifsShare" { "CIFS Share" }
                "DataDomain" { "Data Domain" }
                "ExaGrid" { "ExaGrid" }
                "HPStoreOnce" { "HP StoreOnce" }
                default { "Unknown" }
            }
            $outputObj = Initialize-Object -Name $r.Name -RepoHost $($r.GetHost()).Name.ToLower() -Path $r.Path -Free $r.GetContainer().CachedFreeSpace.InBytes -Total $r.GetContainer().CachedTotalSpace.InBytes -MaxTasks $r.Options.MaxTaskCount -RType $rType
        }
        $outputAry += $outputObj
    }
    End {
        $outputAry
    }
}

Function Get-VBRSORepoInfo {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, ValueFromPipeline = $true)]
        [PSObject[]]$Repository
    )
    Begin {
        $outputAry = @()
        Function Initialize-Object {
            param(
                $Name,
                $Rname,
                $Repohost,
                $Path,
                $Free,
                $Total,
                $Maxtasks,
                $Rtype
            )
            $repoObj = New-Object -TypeName PSObject -Property @{
                SoTarget       = $name
                Target         = $rname
                RepoHost       = $repohost
                Storepath      = $path
                StorageFree    = [Math]::Round([Decimal]$free / 1GB, 2)
                StorageTotal   = [Math]::Round([Decimal]$total / 1GB, 2)
                FreePercentage = [Math]::Round(($free / $total) * 100)
                MaxTasks       = $maxtasks
                rType          = $rtype
            }
            Return $repoObj
        }
    }
    Process {
        Foreach ($rs in $Repository) {
            ForEach ($rp in $rs.Extent) {
                $r = $rp.Repository
                # Refresh Repository Size Info
                [Veeam.Backup.Core.CBackupRepositoryEx]::SyncSpaceInfoToDb($r, $true)
                $rType = switch ($r.Type) {
                    "WinLocal" { "Windows Local" }
                    "LinuxLocal" { "Linux Local" }
                    "CifsShare" { "CIFS Share" }
                    "DataDomain" { "Data Domain" }
                    "ExaGrid" { "ExaGrid" }
                    "HPStoreOnce" { "HP StoreOnce" }
                    default { "Unknown" }
                }
                $outputObj = Initialize-Object -Name $rs.Name -Rname $r.Name -Repohost $($r.GetHost()).Name.ToLower() -Path $r.Path -Free $r.GetContainer().CachedFreeSpace.InBytes -Total $r.GetContainer().CachedTotalSpace.InBytes -Maxtasks $r.Options.MaxTaskCount -Rtype $rType
                $outputAry += $outputObj
            }
        }
    }
    End {
        $outputAry
    }
}

function Get-RepoPermission {
    $outputAry = @()
    $repoEPPerms = $script:repoList | get-vbreppermission
    $repoEPPermsSo = $script:repoListSo | get-vbreppermission
    ForEach ($repo in $repoEPPerms) {
        $objoutput = New-Object -TypeName PSObject -Property @{
            Name                 = (Get-VBRBackupRepository | Where-Object { $_.Id -eq $repo.RepositoryId }).Name
            "Permission Type"    = $repo.PermissionType
            Users                = $repo.Users | Out-String
            "Encryption Enabled" = $repo.IsEncryptionEnabled
        }
        $outputAry += $objoutput
    }
    ForEach ($repo in $repoEPPermsSo) {
        $objoutput = New-Object -TypeName PSObject -Property @{
            Name                 = "[SO] $((Get-VBRBackupRepository -ScaleOut | Where-Object {$_.Id -eq $repo.RepositoryId}).Name)"
            "Permission Type"    = $repo.PermissionType
            Users                = $repo.Users | Out-String
            "Encryption Enabled" = $repo.IsEncryptionEnabled
        }
        $outputAry += $objoutput
    }
    $outputAry
}

Function Get-VBRReplicaTarget {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [PSObject[]]$InputObject
    )
    BEGIN {
        $outputAry = @()
        $dsAry = @()
        If (($null -ne $Name) -and ($null -eq $InputObject)) {
            $InputObject = Get-VBRJob -Name $Name
        }
    }
    PROCESS {
        Foreach ($obj in $InputObject) {
            If (($dsAry -contains $obj.ViReplicaTargetOptions.DatastoreName) -eq $false) {
                $esxi = $obj.GetTargetHost()
                $dtstr = $esxi | Find-VBRViDatastore -Name $obj.ViReplicaTargetOptions.DatastoreName
                $objoutput = New-Object -TypeName PSObject -Property @{
                    Target         = $esxi.Name
                    Datastore      = $obj.ViReplicaTargetOptions.DatastoreName
                    StorageFree    = [Math]::Round([Decimal]$dtstr.FreeSpace / 1GB, 2)
                    StorageTotal   = [Math]::Round([Decimal]$dtstr.Capacity / 1GB, 2)
                    FreePercentage = [Math]::Round(($dtstr.FreeSpace / $dtstr.Capacity) * 100)
                }
                $dsAry = $dsAry + $obj.ViReplicaTargetOptions.DatastoreName
                $outputAry = $outputAry + $objoutput
            } Else {
                return
            }
        }
    }
    END {
        $outputAry | Select-Object Target, Datastore, StorageFree, StorageTotal, FreePercentage
    }
}

Function Get-VeeamSupportDate {
    # Query for license info
    $licenseInfo = Get-VBRInstalledLicense

    $type = $licenseinfo.Type

    switch ( $type ) {
        'Perpetual' {
            $date = $licenseInfo.SupportExpirationDate
        }
        'Evaluation' {
            # No expiration
            $date = Get-Date
        }
        'Subscription' {
            $date = $licenseInfo.ExpirationDate
        }
        'Rental' {
            $date = $licenseInfo.ExpirationDate
        }
        'NFR' {
            $date = $licenseInfo.ExpirationDate
        }

    }

    [PSCustomObject]@{
       LicType    = $type
       ExpDate    = $date.ToShortDateString()
       DaysRemain = ($date - (Get-Date)).Days
    }
}

Function Get-VeeamWinServer {
    $vservers = @{}
    $outputAry = @()
    $vservers.add($($script:vbrServerObj.Name), "VBRServer")
    Foreach ($srv in $script:proxyList) {
        If (!$vservers.ContainsKey($srv.Host.Name)) {
            $vservers.Add($srv.Host.Name, "ProxyServer")
        }
    }
    Foreach ($srv in $script:repoList) {
        If ($srv.Type -ne "LinuxLocal" -and !$vservers.ContainsKey($srv.gethost().Name)) {
            $vservers.Add($srv.gethost().Name, "RepoServer")
        }
    }
    Foreach ($rs in $script:repoListSo) {
        ForEach ($rp in $rs.Extent) {
            $r = $rp.Repository
            $rName = $($r.GetHost()).Name
            If ($r.Type -ne "LinuxLocal" -and !$vservers.ContainsKey($rName)) {
                $vservers.Add($rName, "RepoSoServer")
            }
        }
    }
    Foreach ($srv in $script:tapesrvList) {
        If (!$vservers.ContainsKey($srv.Name)) {
            $vservers.Add($srv.Name, "TapeServer")
        }
    }
    $vservers = $vservers.GetEnumerator() | Sort-Object Name
    Foreach ($vserver in $vservers) {
        $outputAry += $vserver.Name
    }
    return $outputAry
}

Function Get-VeeamService {
    param (
        [PSObject]$InputObject
    )
    $outputAry = @()
    Foreach ($obj in $InputObject) {
        $output = @()
        Try {
            $output = Get-Service -computername $obj -Name "*Veeam*" -exclude "SQLAgent*" |
            Select-Object @{Name = "Server Name"; Expression = { $obj.ToLower() } }, @{Name = "Service Name"; Expression = { $_.DisplayName } }, Status
        } Catch {
            $output = New-Object PSObject -Property @{
                "Server Name"  = $obj.ToLower()
                "Service Name" = "Unable to connect"
                Status         = "Unknown"
            }
        }
        $outputAry += $output
    }
    $outputAry
}

Function Get-VMsBackupStatus {
    $outputary = @()
    # Convert exclusion list to simple regular expression
    $excludevms_regex = ('(?i)^(' + (($script:excludeVMs | ForEach-Object { [regex]::escape($_) }) -join "|") + ')$') -replace "\\\*", ".*"
    $excludefolder_regex = ('(?i)^(' + (($script:excludeFolder | ForEach-Object { [regex]::escape($_) }) -join "|") + ')$') -replace "\\\*", ".*"
    $excludedc_regex = ('(?i)^(' + (($script:excludeDC | ForEach-Object { [regex]::escape($_) }) -join "|") + ')$') -replace "\\\*", ".*"
    $excludecluster_regex = ('(?i)^(' + (($script:excludeCluster | ForEach { [regex]::escape($_) }) -join "|") + ')$') -replace "\\\*", ".*"
    $vms = @{}
    # Build a hash table of all VMs.  Key is either Job Object Id (for any VM ever in a Veeam job) or vCenter ID+MoRef
    # Assume unprotected (!), and populate Cluster, DataCenter, and Name fields for hash key value
    Find-VBRViEntity |
    Where-Object { $_.Type -eq "Vm" -and $_.VmFolderName -notmatch $excludefolder_regex } |
    Where-Object { $_.Name -notmatch $excludevms_regex } |
    Where-Object { $_.Path.Split("\")[1] -notmatch $excludedc_regex } |
    Where-Object { $_.Path.Split("\")[2] -notmatch $excludecluster_regex } |
    ForEach-Object {
        $key = @(($_.FindObject().Id, $_.Id) | Where-Object { $null -ne $_ })[0]
        $vms.Add($key, @("!", $_.Path.Split("\")[0], $_.Path.Split("\")[1], $_.Path.Split("\")[2], $_.Name, "1/11/1911", "1/11/1911", "", $_.VmFolderName))
    }
    If (!$script:excludeTemp) {
        Find-VBRViEntity -VMsandTemplates |
        Where-Object { $_.Type -eq "Vm" -and $_.IsTemplate -eq "True" -and $_.VmFolderName -notmatch $excludefolder_regex } |
        Where-Object { $_.Name -notmatch $excludevms_regex } |
        Where-Object { $_.Path.Split("\")[1] -notmatch $excludedc_regex } |
        Where-Object { $_.VmHostName -notmatch $excludecluster_regex } |
        ForEach-Object {
            $key = @(($_.FindObject().Id, $_.Id) | Where-Object { $null -ne $_ })[0]
            $vms.Add($key, @("!", $_.Path.Split("\")[0], $_.Path.Split("\")[1], $_.VmHostName, "[template] $($_.Name)", "1/11/1911", "1/11/1911", "", $_.VmFolderName))
        }
    }
    # Find all backup task sessions that have ended in the last x hours
    $vbrtasksessions = (Get-VBRBackupSession |
        Where-Object { ($_.JobType -eq "Backup") -and ($_.EndTime -ge (Get-Date).addhours(-$script:HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$script:HourstoCheck) -or $_.State -eq "Working") }) |
    Get-VBRTaskSession | Where-Object { $_.Status -notmatch "InProgress|Pending" }
    # Compare VM list to session list and update found VMs status
    If ($vbrtasksessions) {
        Foreach ($vmtask in $vbrtasksessions) {
            If ($vms.ContainsKey($vmtask.Info.ObjectId)) {
                If ((Get-Date $vmtask.Progress.StartTimeLocal) -ge (Get-Date $vms[$vmtask.Info.ObjectId][5])) {
                    If ($vmtask.Status -eq "Success") {
                        $vms[$vmtask.Info.ObjectId][0] = $vmtask.Status
                        $vms[$vmtask.Info.ObjectId][5] = $vmtask.Progress.StartTimeLocal
                        $vms[$vmtask.Info.ObjectId][6] = $vmtask.Progress.StopTimeLocal
                        $vms[$vmtask.Info.ObjectId][7] = ""
                    } ElseIf ($vms[$vmtask.Info.ObjectId][0] -ne "Success") {
                        $vms[$vmtask.Info.ObjectId][0] = $vmtask.Status
                        $vms[$vmtask.Info.ObjectId][5] = $vmtask.Progress.StartTimeLocal
                        $vms[$vmtask.Info.ObjectId][6] = $vmtask.Progress.StopTimeLocal
                        $vms[$vmtask.Info.ObjectId][7] = ($vmtask.GetDetails()).Replace("<br />", "ZZbrZZ")
                    }
                } ElseIf ($vms[$vmtask.Info.ObjectId][0] -match "Warning|Failed" -and $vmtask.Status -eq "Success") {
                    $vms[$vmtask.Info.ObjectId][0] = $vmtask.Status
                    $vms[$vmtask.Info.ObjectId][5] = $vmtask.Progress.StartTimeLocal
                    $vms[$vmtask.Info.ObjectId][6] = $vmtask.Progress.StopTimeLocal
                    $vms[$vmtask.Info.ObjectId][7] = ""
                }
            }
        }
    }
    Foreach ($vm in $vms.GetEnumerator()) {
        $objoutput = New-Object -TypeName PSObject -Property @{
            Status     = $vm.Value[0]
            Name       = $vm.Value[4]
            vCenter    = $vm.Value[1]
            Datacenter = $vm.Value[2]
            Cluster    = $vm.Value[3]
            StartTime  = $vm.Value[5]
            StopTime   = $vm.Value[6]
            Details    = $vm.Value[7]
            Folder     = $vm.Value[8]
        }
        $outputAry += $objoutput
    }
    $outputAry
}

function Get-Duration {
    param ($ts)
    $days = ""
    If ($ts.Days -gt 0) {
        $days = "{0}:" -f $ts.Days
    }
    "{0}{1}:{2,2:D2}:{3,2:D2}" -f $days, $ts.Hours, $ts.Minutes, $ts.Seconds
}

function Get-BackupSize {
    param ($backups)
    $outputObj = @()
    Foreach ($backup in $backups) {
        $backupSize = 0
        $dataSize = 0
        $files = $backup.GetAllStorages()
        Foreach ($file in $Files) {
            $backupSize += [math]::Round([long]$file.Stats.BackupSize / 1GB, 2)
            $dataSize += [math]::Round([long]$file.Stats.DataSize / 1GB, 2)
        }
        $repo = If ($($script:repoList | Where-Object { $_.Id -eq $backup.RepositoryId }).Name) {
            $($script:repoList | Where-Object { $_.Id -eq $backup.RepositoryId }).Name
        } Else {
            $($script:repoListSo | Where-Object { $_.Id -eq $backup.RepositoryId }).Name
        }
        $vbrMasterHash = @{
            JobName    = $backup.JobName
            VMCount    = $backup.VmCount
            Repo       = $repo
            DataSize   = $dataSize
            BackupSize = $backupSize
        }
        $vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
        $outputObj += $vbrMasterObj
    }
    $outputObj
}
Function Get-MultiJob {
    $outputAry = @()
    $vmMultiJobs = (Get-VBRBackupSession |
        Where-Object { ($_.JobType -eq "Backup") -and ($_.EndTime -ge (Get-Date).addhours(-$script:HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$script:HourstoCheck) -or $_.State -eq "Working") }) |
    Get-VBRTaskSession | Select-Object Name, @{Name = "VMID"; Expression = { $_.Info.ObjectId } }, JobName -Unique | Group-Object Name, VMID | Where-Object { $_.Count -gt 1 } | Select-Object -ExpandProperty Group
    ForEach ($vm in $vmMultiJobs) {
        $objID = $vm.VMID
        $viEntity = Find-VBRViEntity -name $vm.Name | Where-Object { $_.FindObject().Id -eq $objID }
        If ($null -ne $viEntity) {
            $objoutput = New-Object -TypeName PSObject -Property @{
                Name       = $vm.Name
                vCenter    = $viEntity.Path.Split("\")[0]
                Datacenter = $viEntity.Path.Split("\")[1]
                Cluster    = $viEntity.Path.Split("\")[2]
                Folder     = $viEntity.VMFolderName
                JobName    = $vm.JobName
            }
            $outputAry += $objoutput
        } Else {
            #assume Template
            $viEntity = Find-VBRViEntity -VMsAndTemplates -name $vm.Name | Where-Object { $_.FindObject().Id -eq $objID }
            If ($null -ne $viEntity) {
                $objoutput = New-Object -TypeName PSObject -Property @{
                    Name       = "[template] " + $vm.Name
                    vCenter    = $viEntity.Path.Split("\")[0]
                    Datacenter = $viEntity.Path.Split("\")[1]
                    Cluster    = $viEntity.VmHostName
                    Folder     = $viEntity.VMFolderName
                    JobName    = $vm.JobName
                }
            }
            If ($objoutput) {
                $outputAry += $objoutput
            }
        }
    }
    $outputAry
}
#endregion

#region Report
# HTML Stuff
$headerObj = @"
<html>
        <head>
                <title>$rptTitle</title>
                        <style>
                            body {font-family: Tahoma; background-color:#ffffff;}
                            table {font-family: Tahoma;width: $($rptWidth)%;font-size: 12px;border-collapse:collapse;}
                            <!-- table tr:nth-child(odd) td {background: #e2e2e2;} -->
                            th {background-color: #e2e2e2;border: 1px solid #a7a9ac;border-bottom: none;}
                            td {background-color: #ffffff;border: 1px solid #a7a9ac;padding: 2px 3px 2px 3px;}
                        </style>
        </head>
"@

$bodyTop = @"
        <body>
                <div>
                    <img src="file:///F:/Report/veeam-logo.png" style="width: 112px; height: 20px;">
                </div>
                <center>
                        <table>
                            <tr>
                                <td style="width: 50%;height: 24px;border: none;background-color: ZZhdbgZZ;color: Black;font-size: 24px;vertical-align: bottom;text-align: left;padding: 30px 0px 30px 15px;">$rptTitle</td>
                            </tr>
                        </table>
                        <table>
                                <tr>
                                        <td style="height: 2px;background-color: #626365;padding: 5px 0 0 7px;border-top: 5px solid white;border-bottom: none;"></td>
                                </tr>
                                
                                <tr>
                                        <td style="width: 50%;height: 24px;border: none;background-color: ZZhdbgZZ;color: Black;font-size: 12px;vertical-align: bottom;text-align: left;padding: 0px 25px 10px 0px;">Description:</td>
                                </tr>
              
                                <tr>
                                        <td style="width: 50%;height: 14px;border: none;background-color: ZZhdbgZZ;color: Black;font-size: 12px;vertical-align: bottom;text-align: left;padding: 10px 5px 2px 0px;">Report Date: $(Get-Date -format g)</td>
                                </tr>
                                <tr>
                                        <td style="width: 50%;height: 14px;border: none;background-color: ZZhdbgZZ;color: Black;font-size: 12px;vertical-align: bottom;text-align: left;padding: 0px 5px 2px 0px;">Veeam Backup Server: $vbrServer</td>
                                </tr>
                                <tr>
                                        <td style="width: 50%;height: 14px;border: none;background-color: ZZhdbgZZ;color: Black;font-size: 12px;vertical-align: bottom;text-align: left;padding: 0px 5px 2px 0px;">Veeam Backup Server Version: v$VeeamVersion</td>
                                </tr>
                                <tr>
                                        <td style="width: 50%;height: 14px;border: none;background-color: ZZhdbgZZ;color: Black;font-size: 12px;vertical-align: bottom;text-align: left;padding: 0px 5px 10px 0px;">$rptMode</td>
                                </tr>
                                <tr>
                                        <td style="height: 2px;background-color: #626365;padding: 5px 0 0 7px;border-top: 5px solid white;border-bottom: none;"></td>
                                </tr>
                        </table>
"@

$subHead01 = @"
<table>
                                <tr>
                                        <td style="height: 35px;background-color: #f3f4f4;color: #626365;font-size: 16px;padding: 5px 0 0 15px;border-top: 5px solid white;border-bottom: none;">
"@

$subHead01suc = @"
<table>
                                 <tr>
                                        <td style="height: 35px;background-color: #00b050;color: #ffffff;font-size: 16px;padding: 5px 0 0 15px;border-top: 5px solid white;border-bottom: none;">
"@

$subHead01war = @"
<table>
                                 <tr>
                                        <td style="height: 35px;background-color: #ffd96c;color: #ffffff;font-size: 16px;padding: 5px 0 0 15px;border-top: 5px solid white;border-bottom: none;">
"@

$subHead01err = @"
<table>
                                <tr>
                                        <td style="height: 35px;background-color: #ff6d6c;color: #ffffff;font-size: 16px;padding: 5px 0 0 15px;border-top: 5px solid white;border-bottom: none;">
"@

$subHead02 = @"
</td>
                                </tr>
                         </table>
"@

$HTMLbreak = @"
<table>
                                <tr>
                                        <td style="height: 10px;background-color: #626365;padding: 5px 0 0 15px;border-top: 5px solid white;border-bottom: none;"></td>
                                </tr>
                        </table>
"@

$footerObj = @"
<table>
                                <tr>
                                        <td style="height: 15px;background-color: #ffffff;border: none;color: #626365;font-size: 10px;text-align:center;">Veeam Report Platform maintained by Phelipe Garcia <a href="https://github.com/phelipegarcia" target="_blank"></td>
                                </tr>
                        </table>
                </center>
        </body>
</html>
"@

#Get VM Backup Status
$vmStatus = @()
If ($showSummaryProtect + $showUnprotectedVMs + $showProtectedVMs) {
    $vmStatus = Get-VMsBackupStatus
}
# VMs Missing Backups
$missingVMs = @($vmStatus | Where-Object { $_.Status -match "!|Failed" })
ForEach ($VM in $missingVMs) {
    If ($VM.Status -eq "!") {
        $VM.Details = "No Backup Task has completed"
        $VM.StartTime = ""
        $VM.StopTime = ""
    }
}
# VMs Successfuly Backed Up
$successVMs = @($vmStatus | Where-Object { $_.Status -eq "Success" })
# VMs Backed Up w/Warning
$warnVMs = @($vmStatus | Where-Object { $_.Status -eq "Warning" })

# Get VM Backup Protection Summary
$bodySummaryProtect = $null
$sumprotectHead = $subHead01
If ($showSummaryProtect) {
    If (@($successVMs).Count -ge 1) {
        $percentProt = 1
        $sumprotectHead = $subHead01suc
    }
    If (@($warnVMs).Count -ge 1) {
        $percentWarn = "*"
        $sumprotectHead = $subHead01war
    } Else {
        $percentWarn = ""
    }
    If (@($missingVMs).Count -ge 1) {
        $percentProt = (@($warnVMs).Count + @($successVMs).Count) / (@($warnVMs).Count + @($successVMs).Count + @($missingVMs).Count)
        $sumprotectHead = $subHead01err
    }
    $vbrMasterHash = @{
        WarningVM   = @($warnVMs).Count
        ProtectedVM = @($successVMs).Count
        FailedVM    = @($missingVMs).Count
        PercentProt = "{0:P2}{1}" -f $percentProt, $percentWarn

    }
    $vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
    $summaryProtect = $vbrMasterObj | Select-Object @{Name = "% Protected"; Expression = { $_.PercentProt } },
    @{Name = "Fully Protected VMs"; Expression = { $_.ProtectedVM } },
    @{Name = "Protected VMs w/Warnings"; Expression = { $_.WarningVM } },
    @{Name = "Unprotected VMs"; Expression = { $_.FailedVM } }
    $bodySummaryProtect = $summaryProtect | ConvertTo-HTML -Fragment
    $bodySummaryProtect = $sumprotectHead + "VM Backup Protection Summary" + $subHead02 + $bodySummaryProtect
}

# Get VMs Missing Backups
$bodyMissing = $null
If ($showUnprotectedVMs) {
    If ($null -ne $missingVMs) {
        $missingVMs = $missingVMs | Sort-Object vCenter, Datacenter, Cluster, Name | Select-Object Name, vCenter, Datacenter, Cluster, Folder,
        @{Name = "Last Start Time"; Expression = { $_.StartTime } }, @{Name = "Last End Time"; Expression = { $_.StopTime } }, Details | ConvertTo-HTML -Fragment
        $bodyMissing = $subHead01err + "VMs with No Successful Backups within RPO" + $subHead02 + $missingVMs
    }
}

# Get VMs Backed Up w/Warnings
$bodyWarning = $null
If ($showProtectedVMs) {
    If ($null -ne $warnVMs) {
        $warnVMs = $warnVMs | Sort-Object vCenter, Datacenter, Cluster, Name | Select-Object Name, vCenter, Datacenter, Cluster, Folder,
        @{Name = "Last Start Time"; Expression = { $_.StartTime } }, @{Name = "Last End Time"; Expression = { $_.StopTime } }, Details | ConvertTo-HTML -Fragment
        $bodyWarning = $subHead01war + "VMs with only Backups with Warnings within RPO" + $subHead02 + $warnVMs
    }
}

# Get VMs Successfuly Backed Up
$bodySuccess = $null
If ($showProtectedVMs) {
    If ($null -ne $successVMs) {
        $successVMs = $successVMs | Sort-Object vCenter, Datacenter, Cluster, Name | Select-Object Name, vCenter, Datacenter, Cluster, Folder,
        @{Name = "Last Start Time"; Expression = { $_.StartTime } }, @{Name = "Last End Time"; Expression = { $_.StopTime } } | ConvertTo-HTML -Fragment
        $bodySuccess = $subHead01suc + "VMs with Successful Backups within RPO" + $subHead02 + $successVMs
    }
}

# Get VMs Backed Up by Multiple Jobs
$bodyMultiJobs = $null
If ($showMultiJobs) {
    $multiJobs = @(Get-MultiJob)
    If ($multiJobs.Count -gt 0) {
        $bodyMultiJobs = $multiJobs | Sort-Object vCenter, Datacenter, Cluster, Name | Select-Object Name, vCenter, Datacenter, Cluster, Folder,
        @{Name = "Job Name"; Expression = { $_.JobName } } | ConvertTo-HTML -Fragment
        $bodyMultiJobs = $subHead01err + "VMs Backed Up by Multiple Jobs within RPO" + $subHead02 + $bodyMultiJobs
    }
}

# Get Backup Summary Info
$bodySummaryBk = $null
If ($showSummaryBk) {
    $vbrMasterHash = @{
        "Failed"      = @($failedSessionsBk).Count
        "Sessions"    = If ($sessListBk) { @($sessListBk).Count } Else { 0 }
        "Read"        = $totalReadBk
        "Transferred" = $totalXferBk
        "Successful"  = @($successSessionsBk).Count
        "Warning"     = @($warningSessionsBk).Count
        "Fails"       = @($failsSessionsBk).Count
        "Running"     = @($runningSessionsBk).Count
    }
    $vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
    If ($onlyLastBk) {
        $total = "Jobs Run"
    } Else {
        $total = "Total Sessions"
    }
    $arrSummaryBk = $vbrMasterObj | Select-Object @{Name = $total; Expression = { $_.Sessions } },
    @{Name = "Read (GB)"; Expression = { $_.Read } }, @{Name = "Transferred (GB)"; Expression = { $_.Transferred } },
    @{Name = "Running"; Expression = { $_.Running } }, @{Name = "Successful"; Expression = { $_.Successful } },
    @{Name = "Warnings"; Expression = { $_.Warning } }, @{Name = "Failures"; Expression = { $_.Fails } },
    @{Name = "Failed"; Expression = { $_.Failed } }
    $bodySummaryBk = $arrSummaryBk | ConvertTo-HTML -Fragment
    If ($arrSummaryBk.Failed -gt 0) {
        $summaryBkHead = $subHead01err
    } ElseIf ($arrSummaryBk.Warnings -gt 0) {
        $summaryBkHead = $subHead01war
    } ElseIf ($arrSummaryBk.Successful -gt 0) {
        $summaryBkHead = $subHead01suc
    } Else {
        $summaryBkHead = $subHead01
    }
    $bodySummaryBk = $summaryBkHead + "Backup Results Summary" + $subHead02 + $bodySummaryBk
}

# Get Backup Job Status
$bodyJobsBk = $null
If ($showJobsBk) {
    If ($allJobsBk.count -gt 0) {
        $bodyJobsBk = @()
        Foreach ($bkJob in $allJobsBk) {
            $bodyJobsBk += $bkJob | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
            @{Name = "Enabled"; Expression = { $_.IsScheduleEnabled } },
            @{Name = "Status"; Expression = {
                    If ($bkJob.IsRunning) {
                        $currentSess = $runningSessionsBk | Where-Object { $_.JobName -eq $bkJob.Name }
                        $csessPercent = $currentSess.Progress.Percents
                        $csessSpeed = [Math]::Round($currentSess.Progress.AvgSpeed / 1MB, 2)
                        $cStatus = "$($csessPercent)% completed at $($csessSpeed) MB/s"
                        $cStatus
                    } Else {
                        "Stopped"
                    }
                }
            },
            @{Name = "Target Repo"; Expression = {
                    If ($($repoList | Where-Object { $_.Id -eq $BkJob.Info.TargetRepositoryId }).Name) {
                        $($repoList | Where-Object { $_.Id -eq $BkJob.Info.TargetRepositoryId }).Name
                    } Else {
                        $($repoListSo | Where-Object { $_.Id -eq $BkJob.Info.TargetRepositoryId }).Name
                    }
                }
            },
            @{Name = "Next Run"; Expression = {
                    If ($_.IsScheduleEnabled -eq $false) { "<Disabled>" }
                    ElseIf ($_.Options.JobOptions.RunManually) { "<not scheduled>" }
                    ElseIf ($_.ScheduleOptions.IsContinious) { "<Continious>" }
                    ElseIf ($_.ScheduleOptions.OptionsScheduleAfterJob.IsEnabled) { "After [" + $(($allJobs + $allJobsTp) | Where-Object { $_.Id -eq $bkJob.Info.ParentScheduleId }).Name + "]" }
                    Else { $_.ScheduleOptions.NextRun }
                }
            },
            @{Name = "Last Result"; Expression = { If ($_.Info.LatestStatus -eq "None") { "Unknown" }Else { $_.Info.LatestStatus } } }
        }
        $bodyJobsBk = $bodyJobsBk | Sort-Object "Next Run" | ConvertTo-HTML -Fragment
        $bodyJobsBk = $subHead01 + "Backup Job Status" + $subHead02 + $bodyJobsBk
    }
}

# Get Backup Job Size
$bodyJobSizeBk = $null
If ($showBackupSizeBk) {
    If ($backupsBk.count -gt 0) {
        $bodyJobSizeBk = Get-BackupSize -backups $backupsBk | Sort-Object JobName | Select-Object @{Name = "Job Name"; Expression = { $_.JobName } },
        @{Name = "VM Count"; Expression = { $_.VMCount } },
        @{Name = "Repository"; Expression = { $_.Repo } },
        @{Name = "Data Size (GB)"; Expression = { $_.DataSize } },
        @{Name = "Backup Size (GB)"; Expression = { $_.BackupSize } } | ConvertTo-HTML -Fragment
        $bodyJobSizeBk = $subHead01 + "Backup Job Size" + $subHead02 + $bodyJobSizeBk
    }
}

# Get all Backup Sessions
$bodyAllSessBk = $null
If ($showAllSessBk) {
    If ($sessListBk.count -gt 0) {
        If ($showDetailedBk) {
            $arrAllSessBk = $sessListBk | Sort-Object Creationtime | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
            @{Name = "State"; Expression = { $_.State } },
            @{Name = "Start Time"; Expression = { $_.CreationTime } },
            @{Name = "Stop Time"; Expression = { If ($_.EndTime -eq "1/1/1900 12:00:00 AM") { "-" } Else { $_.EndTime } } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Progress.AvgSpeed / 1MB, 2) } },
            @{Name = "Total (GB)"; Expression = { [Math]::Round($_.Progress.ProcessedSize / 1GB, 2) } },
            @{Name = "Processed (GB)"; Expression = { [Math]::Round($_.Progress.ProcessedUsedSize / 1GB, 2) } },
            @{Name = "Data Read (GB)"; Expression = { [Math]::Round($_.Progress.ReadSize / 1GB, 2) } },
            @{Name = "Transferred (GB)"; Expression = { [Math]::Round($_.Progress.TransferedSize / 1GB, 2) } },
            @{Name = "Dedupe"; Expression = {
                    If ($_.Progress.ReadSize -eq 0) { 0 }
                    Else { ([string][Math]::Round($_.BackupStats.GetDedupeX(), 1)) + "x" } }
            },
            @{Name = "Compression"; Expression = {
                    If ($_.Progress.ReadSize -eq 0) { 0 }
                    Else { ([string][Math]::Round($_.BackupStats.GetCompressX(), 1)) + "x" } }
            },
            @{Name = "Details"; Expression = { ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } }, Result
            $bodyAllSessBk = $arrAllSessBk  | ConvertTo-HTML -Fragment
            If ($arrAllSessBk.Result -match "Failed") {
                $allSessBkHead = $subHead01err
            } ElseIf ($arrAllSessBk.Result -match "Warning") {
                $allSessBkHead = $subHead01war
            } ElseIf ($arrAllSessBk.Result -match "Success") {
                $allSessBkHead = $subHead01suc
            } Else {
                $allSessBkHead = $subHead01
            }
            $bodyAllSessBk = $allSessBkHead + "Backup Sessions" + $subHead02 + $bodyAllSessBk
        } Else {
            $arrAllSessBk = $sessListBk | Sort-Object Creationtime | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
            @{Name = "State"; Expression = { $_.State } },
            @{Name = "Start Time"; Expression = { $_.CreationTime } },
            @{Name = "Stop Time"; Expression = { If ($_.EndTime -eq "1/1/1900 12:00:00 AM") { "-" } Else { $_.EndTime } } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Details"; Expression = { ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } }, Result
            $bodyAllSessBk = $arrAllSessBk | ConvertTo-HTML -Fragment
            If ($arrAllSessBk.Result -match "Failed") {
                $allSessBkHead = $subHead01err
            } ElseIf ($arrAllSessBk.Result -match "Warning") {
                $allSessBkHead = $subHead01war
            } ElseIf ($arrAllSessBk.Result -match "Success") {
                $allSessBkHead = $subHead01suc
            } Else {
                $allSessBkHead = $subHead01
            }
            $bodyAllSessBk = $allSessBkHead + "Backup Sessions" + $subHead02 + $bodyAllSessBk
        }
    }
}

# Get Running Backup Jobs
$bodyRunningBk = $null
If ($showRunningBk) {
    If ($runningSessionsBk.count -gt 0) {
        $bodyRunningBk = $runningSessionsBk | Sort-Object Creationtime | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
        @{Name = "Start Time"; Expression = { $_.CreationTime } },
        @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
        @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Progress.AvgSpeed / 1MB, 2) } },
        @{Name = "Read (GB)"; Expression = { [Math]::Round([Decimal]$_.Progress.ReadSize / 1GB, 2) } },
        @{Name = "Transferred (GB)"; Expression = { [Math]::Round([Decimal]$_.Progress.TransferedSize / 1GB, 2) } },
        @{Name = "% Complete"; Expression = { $_.Progress.Percents } } | ConvertTo-HTML -Fragment
        $bodyRunningBk = $subHead01 + "Running Backup Jobs" + $subHead02 + $bodyRunningBk
    }
}

# Get Backup Sessions with Warnings or Failures
$bodySessWFBk = $null
If ($showWarnFailBk) {
    $sessWF = @($warningSessionsBk + $failsSessionsBk)
    If ($sessWF.count -gt 0) {
        If ($onlyLastBk) {
            $headerWF = "Backup Jobs with Warnings or Failures"
        } Else {
            $headerWF = "Backup Sessions with Warnings or Failures"
        }
        If ($showDetailedBk) {
            $arrSessWFBk = $sessWF | Sort-Object Creationtime | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
            @{Name = "Start Time"; Expression = { $_.CreationTime } },
            @{Name = "Stop Time"; Expression = { $_.EndTime } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Progress.AvgSpeed / 1MB, 2) } },
            @{Name = "Total (GB)"; Expression = { [Math]::Round($_.Progress.ProcessedSize / 1GB, 2) } },
            @{Name = "Processed (GB)"; Expression = { [Math]::Round($_.Progress.ProcessedUsedSize / 1GB, 2) } },
            @{Name = "Data Read (GB)"; Expression = { [Math]::Round($_.Progress.ReadSize / 1GB, 2) } },
            @{Name = "Transferred (GB)"; Expression = { [Math]::Round($_.Progress.TransferedSize / 1GB, 2) } },
            @{Name = "Dedupe"; Expression = {
                    If ($_.Progress.ReadSize -eq 0) { 0 }
                    Else { ([string][Math]::Round($_.BackupStats.GetDedupeX(), 1)) + "x" } }
            },
            @{Name = "Compression"; Expression = {
                    If ($_.Progress.ReadSize -eq 0) { 0 }
                    Else { ([string][Math]::Round($_.BackupStats.GetCompressX(), 1)) + "x" } }
            },
            @{Name = "Details"; Expression = {
                    If ($_.GetDetails() -eq "") { $_ | Get-VBRTaskSession | ForEach-Object { If ($_.GetDetails()) { $_.Name + ": " + ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } } }
                    Else { ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } }
            }, Result
            $bodySessWFBk = $arrSessWFBk | ConvertTo-HTML -Fragment
            If ($arrSessWFBk.Result -match "Failed") {
                $sessWFBkHead = $subHead01err
            } ElseIf ($arrSessWFBk.Result -match "Warning") {
                $sessWFBkHead = $subHead01war
            } ElseIf ($arrSessWFBk.Result -match "Success") {
                $sessWFBkHead = $subHead01suc
            } Else {
                $sessWFBkHead = $subHead01
            }
            $bodySessWFBk = $sessWFBkHead + $headerWF + $subHead02 + $bodySessWFBk
        } Else {
            $arrSessWFBk = $sessWF | Sort-Object Creationtime | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
            @{Name = "Start Time"; Expression = { $_.CreationTime } },
            @{Name = "Stop Time"; Expression = { $_.EndTime } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Details"; Expression = {
                    If ($_.GetDetails() -eq "") { $_ | Get-VBRTaskSession | ForEach-Object { If ($_.GetDetails()) { $_.Name + ": " + ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } } }
                    Else { ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } }
            }, Result
            $bodySessWFBk = $arrSessWFBk | ConvertTo-HTML -Fragment
            If ($arrSessWFBk.Result -match "Failed") {
                $sessWFBkHead = $subHead01err
            } ElseIf ($arrSessWFBk.Result -match "Warning") {
                $sessWFBkHead = $subHead01war
            } ElseIf ($arrSessWFBk.Result -match "Success") {
                $sessWFBkHead = $subHead01suc
            } Else {
                $sessWFBkHead = $subHead01
            }
            $bodySessWFBk = $sessWFBkHead + $headerWF + $subHead02 + $bodySessWFBk
        }
    }
}

# Get Successful Backup Sessions
$bodySessSuccBk = $null
If ($showSuccessBk) {
    If ($successSessionsBk.count -gt 0) {
        If ($onlyLastBk) {
            $headerSucc = "Successful Backup Jobs"
        } Else {
            $headerSucc = "Successful Backup Sessions"
        }
        If ($showDetailedBk) {
            $bodySessSuccBk = $successSessionsBk | Sort-Object Creationtime | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
            @{Name = "Start Time"; Expression = { $_.CreationTime } },
            @{Name = "Stop Time"; Expression = { $_.EndTime } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Progress.AvgSpeed / 1MB, 2) } },
            @{Name = "Total (GB)"; Expression = { [Math]::Round($_.Progress.ProcessedSize / 1GB, 2) } },
            @{Name = "Processed (GB)"; Expression = { [Math]::Round($_.Progress.ProcessedUsedSize / 1GB, 2) } },
            @{Name = "Data Read (GB)"; Expression = { [Math]::Round($_.Progress.ReadSize / 1GB, 2) } },
            @{Name = "Transferred (GB)"; Expression = { [Math]::Round($_.Progress.TransferedSize / 1GB, 2) } },
            @{Name = "Dedupe"; Expression = {
                    If ($_.Progress.ReadSize -eq 0) { 0 }
                    Else { ([string][Math]::Round($_.BackupStats.GetDedupeX(), 1)) + "x" } }
            },
            @{Name = "Compression"; Expression = {
                    If ($_.Progress.ReadSize -eq 0) { 0 }
                    Else { ([string][Math]::Round($_.BackupStats.GetCompressX(), 1)) + "x" } }
            },
            Result  | ConvertTo-HTML -Fragment
            $bodySessSuccBk = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccBk
        } Else {
            $bodySessSuccBk = $successSessionsBk | Sort-Object Creationtime | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
            @{Name = "Start Time"; Expression = { $_.CreationTime } },
            @{Name = "Stop Time"; Expression = { $_.EndTime } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            Result | ConvertTo-HTML -Fragment
            $bodySessSuccBk = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccBk
        }
    }
}

## Gathering tasks after session info has been recorded due to Veeam issue
# Gather all Backup Tasks from Sessions within time frame
$taskListBk = @()
$taskListBk += $sessListBk | Get-VBRTaskSession
$successTasksBk = @($taskListBk | Where-Object { $_.Status -eq "Success" })
$wfTasksBk = @($taskListBk | Where-Object { $_.Status -match "Warning|Failed" })
$runningTasksBk = @()
$runningTasksBk += $runningSessionsBk | Get-VBRTaskSession | Where-Object { $_.Status -match "Pending|InProgress" }

# Get all Backup Tasks
$bodyAllTasksBk = $null
If ($showAllTasksBk) {
    If ($taskListBk.count -gt 0) {
        If ($showDetailedBk) {
            $arrAllTasksBk = $taskListBk | Select-Object @{Name = "VM Name"; Expression = { $_.Name } },
            @{Name = "Job Name"; Expression = { $_.JobSess.Name } },
            @{Name = "Start Time"; Expression = { $_.Progress.StartTimeLocal } },
            @{Name = "Stop Time"; Expression = { If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") { "-" } Else { $_.Progress.StopTimeLocal } } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Progress.AvgSpeed / 1MB, 2) } },
            @{Name = "Total (GB)"; Expression = { [Math]::Round($_.Progress.ProcessedSize / 1GB, 2) } },
            @{Name = "Processed (GB)"; Expression = { [Math]::Round($_.Progress.ProcessedUsedSize / 1GB, 2) } },
            @{Name = "Data Read (GB)"; Expression = { [Math]::Round($_.Progress.ReadSize / 1GB, 2) } },
            @{Name = "Transferred (GB)"; Expression = { [Math]::Round($_.Progress.TransferedSize / 1GB, 2) } },
            @{Name = "Details"; Expression = { ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } }, Status
            $bodyAllTasksBk = $arrAllTasksBk | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
            If ($arrAllTasksBk.Status -match "Failed") {
                $allTasksBkHead = $subHead01err
            } ElseIf ($arrAllTasksBk.Status -match "Warning") {
                $allTasksBkHead = $subHead01war
            } ElseIf ($arrAllTasksBk.Status -match "Success") {
                $allTasksBkHead = $subHead01suc
            } Else {
                $allTasksBkHead = $subHead01
            }
            $bodyAllTasksBk = $allTasksBkHead + "Backup Tasks" + $subHead02 + $bodyAllTasksBk
        } Else {
            $arrAllTasksBk = $taskListBk | Select-Object @{Name = "VM Name"; Expression = { $_.Name } },
            @{Name = "Job Name"; Expression = { $_.JobSess.Name } },
            @{Name = "Start Time"; Expression = { $_.Progress.StartTimeLocal } },
            @{Name = "Stop Time"; Expression = { If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") { "-" } Else { $_.Progress.StopTimeLocal } } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Details"; Expression = { ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } }, Status
            $bodyAllTasksBk = $arrAllTasksBk | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
            If ($arrAllTasksBk.Status -match "Failed") {
                $allTasksBkHead = $subHead01err
            } ElseIf ($arrAllTasksBk.Status -match "Warning") {
                $allTasksBkHead = $subHead01war
            } ElseIf ($arrAllTasksBk.Status -match "Success") {
                $allTasksBkHead = $subHead01suc
            } Else {
                $allTasksBkHead = $subHead01
            }
            $bodyAllTasksBk = $allTasksBkHead + "Backup Tasks" + $subHead02 + $bodyAllTasksBk
        }
    }
}

# Get Running Backup Tasks
$bodyTasksRunningBk = $null
If ($showRunningTasksBk) {
    If ($runningTasksBk.count -gt 0) {
        $bodyTasksRunningBk = $runningTasksBk | Select-Object @{Name = "VM Name"; Expression = { $_.Name } },
        @{Name = "Job Name"; Expression = { $_.JobSess.Name } },
        @{Name = "Start Time"; Expression = { $_.Info.Progress.StartTimeLocal } },
        @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
        @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Progress.AvgSpeed / 1MB, 2) } },
        @{Name = "Read (GB)"; Expression = { [Math]::Round($_.Progress.ReadSize / 1GB, 2) } },
        @{Name = "Transferred (GB)"; Expression = { [Math]::Round($_.Progress.TransferedSize / 1GB, 2) } },
        Status | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
        $bodyTasksRunningBk = $subHead01 + "Running Backup Tasks" + $subHead02 + $bodyTasksRunningBk
    }
}

# Get Backup Tasks with Warnings or Failures
$bodyTaskWFBk = $null
If ($showTaskWFBk) {
    If ($wfTasksBk.count -gt 0) {
        If ($showDetailedBk) {
            $arrTaskWFBk = $wfTasksBk | Select-Object @{Name = "VM Name"; Expression = { $_.Name } },
            @{Name = "Job Name"; Expression = { $_.JobSess.Name } },
            @{Name = "Start Time"; Expression = { $_.Progress.StartTimeLocal } },
            @{Name = "Stop Time"; Expression = { $_.Progress.StopTimeLocal } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Progress.AvgSpeed / 1MB, 2) } },
            @{Name = "Total (GB)"; Expression = { [Math]::Round($_.Progress.ProcessedSize / 1GB, 2) } },
            @{Name = "Processed (GB)"; Expression = { [Math]::Round($_.Progress.ProcessedUsedSize / 1GB, 2) } },
            @{Name = "Data Read (GB)"; Expression = { [Math]::Round($_.Progress.ReadSize / 1GB, 2) } },
            @{Name = "Transferred (GB)"; Expression = { [Math]::Round($_.Progress.TransferedSize / 1GB, 2) } },
            @{Name = "Details"; Expression = { ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } }, Status
            $bodyTaskWFBk = $arrTaskWFBk | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
            If ($arrTaskWFBk.Status -match "Failed") {
                $taskWFBkHead = $subHead01err
            } ElseIf ($arrTaskWFBk.Status -match "Warning") {
                $taskWFBkHead = $subHead01war
            } ElseIf ($arrTaskWFBk.Status -match "Success") {
                $taskWFBkHead = $subHead01suc
            } Else {
                $taskWFBkHead = $subHead01
            }
            $bodyTaskWFBk = $taskWFBkHead + "Backup Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFBk
        } Else {
            $arrTaskWFBk = $wfTasksBk | Select-Object @{Name = "VM Name"; Expression = { $_.Name } },
            @{Name = "Job Name"; Expression = { $_.JobSess.Name } },
            @{Name = "Start Time"; Expression = { $_.Progress.StartTimeLocal } },
            @{Name = "Stop Time"; Expression = { $_.Progress.StopTimeLocal } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Details"; Expression = { ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } }, Status
            $bodyTaskWFBk = $arrTaskWFBk | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
            If ($arrTaskWFBk.Status -match "Failed") {
                $taskWFBkHead = $subHead01err
            } ElseIf ($arrTaskWFBk.Status -match "Warning") {
                $taskWFBkHead = $subHead01war
            } ElseIf ($arrTaskWFBk.Status -match "Success") {
                $taskWFBkHead = $subHead01suc
            } Else {
                $taskWFBkHead = $subHead01
            }
            $bodyTaskWFBk = $taskWFBkHead + "Backup Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFBk
        }
    }
}

# Get Successful Backup Tasks
$bodyTaskSuccBk = $null
If ($showTaskSuccessBk) {
    If ($successTasksBk.count -gt 0) {
        If ($showDetailedBk) {
            $bodyTaskSuccBk = $successTasksBk | Select-Object @{Name = "VM Name"; Expression = { $_.Name } },
            @{Name = "Job Name"; Expression = { $_.JobSess.Name } },
            @{Name = "Start Time"; Expression = { $_.Progress.StartTimeLocal } },
            @{Name = "Stop Time"; Expression = { $_.Progress.StopTimeLocal } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Progress.AvgSpeed / 1MB, 2) } },
            @{Name = "Total (GB)"; Expression = { [Math]::Round($_.Progress.ProcessedSize / 1GB, 2) } },
            @{Name = "Processed (GB)"; Expression = { [Math]::Round($_.Progress.ProcessedUsedSize / 1GB, 2) } },
            @{Name = "Data Read (GB)"; Expression = { [Math]::Round($_.Progress.ReadSize / 1GB, 2) } },
            @{Name = "Transferred (GB)"; Expression = { [Math]::Round($_.Progress.TransferedSize / 1GB, 2) } },
            Status | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
            $bodyTaskSuccBk = $subHead01suc + "Successful Backup Tasks" + $subHead02 + $bodyTaskSuccBk
        } Else {
            $bodyTaskSuccBk = $successTasksBk | Select-Object @{Name = "VM Name"; Expression = { $_.Name } },
            @{Name = "Job Name"; Expression = { $_.JobSess.Name } },
            @{Name = "Start Time"; Expression = { $_.Progress.StartTimeLocal } },
            @{Name = "Stop Time"; Expression = { $_.Progress.StopTimeLocal } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            Status | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
            $bodyTaskSuccBk = $subHead01suc + "Successful Backup Tasks" + $subHead02 + $bodyTaskSuccBk
        }
    }
}

# Get Running VM Restore Sessions
$bodyRestoRunVM = $null
If ($showRestoRunVM) {
    If ($($runningResto).count -gt 0) {
        $bodyRestoRunVM = $runningResto | Sort-Object CreationTime | Select-Object @{Name = "VM Name"; Expression = { $_.Info.VmDisplayName } },
        @{Name = "Restore Type"; Expression = { $_.JobTypeString } }, @{Name = "Start Time"; Expression = { $_.CreationTime } },
        @{Name = "Initiator"; Expression = { $_.Info.Initiator.Name } },
        @{Name = "Reason"; Expression = { $_.Info.Reason } } | ConvertTo-HTML -Fragment
        $bodyRestoRunVM = $subHead01 + "Running VM Restore Sessions" + $subHead02 + $bodyRestoRunVM
    }
}

# Get Completed VM Restore Sessions
$bodyRestoreVM = $null
If ($showRestoreVM) {
    If ($($completeResto).count -gt 0) {
        $arrRestoreVM = $completeResto | Sort-Object CreationTime | Select-Object @{Name = "VM Name"; Expression = { $_.Info.VmDisplayName } },
        @{Name = "Restore Type"; Expression = { $_.JobTypeString } },
        @{Name = "Start Time"; Expression = { $_.CreationTime } }, @{Name = "Stop Time"; Expression = { $_.EndTime } },
        @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime) } },
        @{Name = "Initiator"; Expression = { $_.Info.Initiator.Name } }, @{Name = "Reason"; Expression = { $_.Info.Reason } },
        @{Name = "Result"; Expression = { $_.Info.Result } }
        $bodyRestoreVM = $arrRestoreVM | ConvertTo-HTML -Fragment
        If ($arrRestoreVM.Result -match "Failed") {
            $restoreVMHead = $subHead01err
        } ElseIf ($arrRestoreVM.Result -match "Warning") {
            $restoreVMHead = $subHead01war
        } ElseIf ($arrRestoreVM.Result -match "Success") {
            $restoreVMHead = $subHead01suc
        } Else {
            $restoreVMHead = $subHead01
        }
        $bodyRestoreVM = $restoreVMHead + "Completed VM Restore Sessions" + $subHead02 + $bodyRestoreVM
    }
}

# Get Replication Summary Info
$bodySummaryRp = $null
If ($showSummaryRp) {
    $vbrMasterHash = @{
        "Failed"      = @($failedSessionsRp).Count
        "Sessions"    = If ($sessListRp) { @($sessListRp).Count } Else { 0 }
        "Read"        = $totalReadRp
        "Transferred" = $totalXferRp
        "Successful"  = @($successSessionsRp).Count
        "Warning"     = @($warningSessionsRp).Count
        "Fails"       = @($failsSessionsRp).Count
        "Running"     = @($runningSessionsRp).Count
    }
    $vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
    If ($onlyLastRp) {
        $total = "Jobs Run"
    } Else {
        $total = "Total Sessions"
    }
    $arrSummaryRp = $vbrMasterObj | Select-Object @{Name = $total; Expression = { $_.Sessions } },
    @{Name = "Read (GB)"; Expression = { $_.Read } }, @{Name = "Transferred (GB)"; Expression = { $_.Transferred } },
    @{Name = "Running"; Expression = { $_.Running } }, @{Name = "Successful"; Expression = { $_.Successful } },
    @{Name = "Warnings"; Expression = { $_.Warning } },
    @{Name = "Failed"; Expression = { $_.Failed } }
    $bodySummaryRp = $arrSummaryRp | ConvertTo-HTML -Fragment
    If ($arrSummaryRp.Failed -gt 0) {
        $summaryRpHead = $subHead01err
    } ElseIf ($arrSummaryRp.Warnings -gt 0) {
        $summaryRpHead = $subHead01war
    } ElseIf ($arrSummaryRp.Successful -gt 0) {
        $summaryRpHead = $subHead01suc
    } Else {
        $summaryRpHead = $subHead01
    }
    $bodySummaryRp = $summaryRpHead + "Replication Results Summary" + $subHead02 + $bodySummaryRp
}

# Get Replication Job Status
$bodyJobsRp = $null
If ($showJobsRp) {
    If ($allJobsRp.count -gt 0) {
        $bodyJobsRp = @()
        Foreach ($rpJob in $allJobsRp) {
            $bodyJobsRp += $rpJob | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
            @{Name = "Enabled"; Expression = { $_.Info.IsScheduleEnabled } },
            @{Name = "Status"; Expression = {
                    If ($rpJob.IsRunning) {
                        $currentSess = $runningSessionsRp | Where-Object { $_.JobName -eq $rpJob.Name }
                        $csessPercent = $currentSess.Progress.Percents
                        $csessSpeed = [Math]::Round($currentSess.Info.Progress.AvgSpeed / 1MB, 2)
                        $cStatus = "$($csessPercent)% completed at $($csessSpeed) MB/s"
                        $cStatus
                    } Else {
                        "Stopped"
                    }
                }
            },
            @{Name = "Target"; Expression = { $(Get-VBRServer | Where-Object { $_.Id -eq $rpJob.Info.TargetHostId }).Name } },
            @{Name = "Target Repo"; Expression = {
                    If ($($repoList | Where-Object { $_.Id -eq $rpJob.Info.TargetRepositoryId }).Name) { $($repoList | Where-Object { $_.Id -eq $rpJob.Info.TargetRepositoryId }).Name }
                    Else { $($repoListSo | Where-Object { $_.Id -eq $rpJob.Info.TargetRepositoryId }).Name } }
            },
            @{Name = "Next Run"; Expression = {
                    If ($_.IsScheduleEnabled -eq $false) { "<Disabled>" }
                    ElseIf ($_.Options.JobOptions.RunManually) { "<not scheduled>" }
                    ElseIf ($_.ScheduleOptions.IsContinious) { "<Continious>" }
                    ElseIf ($_.ScheduleOptions.OptionsScheduleAfterJob.IsEnabled) { "After [" + $(($allJobs + $allJobsTp) | Where-Object { $_.Id -eq $rpJob.Info.ParentScheduleId }).Name + "]" }
                    Else { $_.ScheduleOptions.NextRun } }
            },
            @{Name = "Last Result"; Expression = { If ($_.Info.LatestStatus -eq "None") { "" }Else { $_.Info.LatestStatus } } }
        }
        $bodyJobsRp = $bodyJobsRp | Sort-Object "Next Run" | ConvertTo-HTML -Fragment
        $bodyJobsRp = $subHead01 + "Replication Job Status" + $subHead02 + $bodyJobsRp
    }
}

# Get Replication Sessions
$bodyAllSessRp = $null
If ($showAllSessRp) {
    If ($sessListRp.count -gt 0) {
        If ($showDetailedRp) {
            $arrAllSessRp = $sessListRp | Sort-Object Creationtime | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
            @{Name = "State"; Expression = { $_.State } },
            @{Name = "Start Time"; Expression = { $_.CreationTime } },
            @{Name = "Stop Time"; Expression = { If ($_.EndTime -eq "1/1/1900 12:00:00 AM") { "-" } Else { $_.EndTime } } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Info.Progress.AvgSpeed / 1MB, 2) } },
            @{Name = "Total (GB)"; Expression = { [Math]::Round($_.Info.Progress.ProcessedSize / 1GB, 2) } },
            @{Name = "Processed (GB)"; Expression = { [Math]::Round($_.Info.Progress.ProcessedUsedSize / 1GB, 2) } },
            @{Name = "Data Read (GB)"; Expression = { [Math]::Round($_.Info.Progress.ReadSize / 1GB, 2) } },
            @{Name = "Transferred (GB)"; Expression = { [Math]::Round($_.Info.Progress.TransferedSize / 1GB, 2) } },
            @{Name = "Dedupe"; Expression = {
                    If ($_.Progress.ReadSize -eq 0) { 0 }
                    Else { ([string][Math]::Round($_.BackupStats.GetDedupeX(), 1)) + "x" } }
            },
            @{Name = "Compression"; Expression = {
                    If ($_.Progress.ReadSize -eq 0) { 0 }
                    Else { ([string][Math]::Round($_.BackupStats.GetCompressX(), 1)) + "x" } }
            },
            @{Name = "Details"; Expression = { ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } }, Result
            $bodyAllSessRp = $arrAllSessRp | ConvertTo-HTML -Fragment
            If ($arrAllSessRp.Result -match "Failed") {
                $allSessRpHead = $subHead01err
            } ElseIf ($arrAllSessRp.Result -match "Warning") {
                $allSessRpHead = $subHead01war
            } ElseIf ($arrAllSessRp.Result -match "Success") {
                $allSessRpHead = $subHead01suc
            } Else {
                $allSessRpHead = $subHead01
            }
            $bodyAllSessRp = $allSessRpHead + "Replication Sessions" + $subHead02 + $bodyAllSessRp
        } Else {
            $arrAllSessRp = $sessListRp | Sort-Object Creationtime | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
            @{Name = "State"; Expression = { $_.State } },
            @{Name = "Start Time"; Expression = { $_.CreationTime } },
            @{Name = "Stop Time"; Expression = { If ($_.EndTime -eq "1/1/1900 12:00:00 AM") { "-" } Else { $_.EndTime } } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Details"; Expression = { ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } }, Result
            $bodyAllSessRp = $arrAllSessRp | ConvertTo-HTML -Fragment
            If ($arrAllSessRp.Result -match "Failed") {
                $allSessRpHead = $subHead01err
            } ElseIf ($arrAllSessRp.Result -match "Warning") {
                $allSessRpHead = $subHead01war
            } ElseIf ($arrAllSessRp.Result -match "Success") {
                $allSessRpHead = $subHead01suc
            } Else {
                $allSessRpHead = $subHead01
            }
            $bodyAllSessRp = $allSessRpHead + "Replication Sessions" + $subHead02 + $bodyAllSessRp
        }
    }
}

# Get Running Replication Jobs
$bodyRunningRp = $null
If ($showRunningRp) {
    If ($runningSessionsRp.count -gt 0) {
        $bodyRunningRp = $runningSessionsRp | Sort-Object Creationtime | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
        @{Name = "Start Time"; Expression = { $_.CreationTime } },
        @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
        @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Info.Progress.AvgSpeed / 1MB, 2) } },
        @{Name = "Read (GB)"; Expression = { [Math]::Round([Decimal]$_.Progress.ReadSize / 1GB, 2) } },
        @{Name = "Transferred (GB)"; Expression = { [Math]::Round([Decimal]$_.Progress.TransferedSize / 1GB, 2) } },
        @{Name = "% Complete"; Expression = { $_.Progress.Percents } } | ConvertTo-HTML -Fragment
        $bodyRunningRp = $subHead01 + "Running Replication Jobs" + $subHead02 + $bodyRunningRp
    }
}

# Get Replication Sessions with Warnings or Failures
$bodySessWFRp = $null
If ($showWarnFailRp) {
    $sessWF = @($warningSessionsRp + $failsSessionsRp)
    If ($sessWF.count -gt 0) {
        If ($onlyLastRp) {
            $headerWF = "Replication Jobs with Warnings or Failures"
        } Else {
            $headerWF = "Replication Sessions with Warnings or Failures"
        }
        If ($showDetailedRp) {
            $arrSessWFRp = $sessWF | Sort-Object Creationtime | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
            @{Name = "Start Time"; Expression = { $_.CreationTime } },
            @{Name = "Stop Time"; Expression = { $_.EndTime } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Info.Progress.AvgSpeed / 1MB, 2) } },
            @{Name = "Total (GB)"; Expression = { [Math]::Round($_.Info.Progress.ProcessedSize / 1GB, 2) } },
            @{Name = "Processed (GB)"; Expression = { [Math]::Round($_.Info.Progress.ProcessedUsedSize / 1GB, 2) } },
            @{Name = "Data Read (GB)"; Expression = { [Math]::Round($_.Info.Progress.ReadSize / 1GB, 2) } },
            @{Name = "Transferred (GB)"; Expression = { [Math]::Round($_.Info.Progress.TransferedSize / 1GB, 2) } },
            @{Name = "Dedupe"; Expression = {
                    If ($_.Progress.ReadSize -eq 0) { 0 }
                    Else { ([string][Math]::Round($_.BackupStats.GetDedupeX(), 1)) + "x" } }
            },
            @{Name = "Compression"; Expression = {
                    If ($_.Progress.ReadSize -eq 0) { 0 }
                    Else { ([string][Math]::Round($_.BackupStats.GetCompressX(), 1)) + "x" } }
            },
            @{Name = "Details"; Expression = {
                    If ($_.GetDetails() -eq "") { $_ | Get-VBRTaskSession | ForEach-Object { If ($_.GetDetails()) { $_.Name + ": " + ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } } }
                    Else { ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } }
            }, Result
            $bodySessWFRp = $arrSessWFRp | ConvertTo-HTML -Fragment
            If ($arrSessWFRp.Result -match "Failed") {
                $sessWFRpHead = $subHead01err
            } ElseIf ($arrSessWFRp.Result -match "Warning") {
                $sessWFRpHead = $subHead01war
            } ElseIf ($arrSessWFRp.Result -match "Success") {
                $sessWFRpHead = $subHead01suc
            } Else {
                $sessWFRpHead = $subHead01
            }
            $bodySessWFRp = $sessWFRpHead + $headerWF + $subHead02 + $bodySessWFRp
        } Else {
            $arrSessWFRp = $sessWF | Sort-Object Creationtime | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
            @{Name = "Start Time"; Expression = { $_.CreationTime } },
            @{Name = "Stop Time"; Expression = { $_.EndTime } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Details"; Expression = {
                    If ($_.GetDetails() -eq "") { $_ | Get-VBRTaskSession | ForEach-Object { If ($_.GetDetails()) { $_.Name + ": " + ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } } }
                    Else { ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } }
            }, Result
            $bodySessWFRp = $arrSessWFRp | ConvertTo-HTML -Fragment
            If ($arrSessWFRp.Result -match "Failed") {
                $sessWFRpHead = $subHead01err
            } ElseIf ($arrSessWFRp.Result -match "Warning") {
                $sessWFRpHead = $subHead01war
            } ElseIf ($arrSessWFRp.Result -match "Success") {
                $sessWFRpHead = $subHead01suc
            } Else {
                $sessWFRpHead = $subHead01
            }
            $bodySessWFRp = $sessWFRpHead + $headerWF + $subHead02 + $bodySessWFRp
        }
    }
}

# Get Successful Replication Sessions
$bodySessSuccRp = $null
If ($showSuccessRp) {
    If ($successSessionsRp.count -gt 0) {
        If ($onlyLastRp) {
            $headerSucc = "Successful Replication Jobs"
        } Else {
            $headerSucc = "Successful Replication Sessions"
        }
        If ($showDetailedRp) {
            $bodySessSuccRp = $successSessionsRp | Sort-Object Creationtime | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
            @{Name = "Start Time"; Expression = { $_.CreationTime } },
            @{Name = "Stop Time"; Expression = { $_.EndTime } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Info.Progress.AvgSpeed / 1MB, 2) } },
            @{Name = "Total (GB)"; Expression = { [Math]::Round($_.Info.Progress.ProcessedSize / 1GB, 2) } },
            @{Name = "Processed (GB)"; Expression = { [Math]::Round($_.Info.Progress.ProcessedUsedSize / 1GB, 2) } },
            @{Name = "Data Read (GB)"; Expression = { [Math]::Round($_.Info.Progress.ReadSize / 1GB, 2) } },
            @{Name = "Transferred (GB)"; Expression = { [Math]::Round($_.Info.Progress.TransferedSize / 1GB, 2) } },
            @{Name = "Dedupe"; Expression = {
                    If ($_.Progress.ReadSize -eq 0) { 0 }
                    Else { ([string][Math]::Round($_.BackupStats.GetDedupeX(), 1)) + "x" } }
            },
            @{Name = "Compression"; Expression = {
                    If ($_.Progress.ReadSize -eq 0) { 0 }
                    Else { ([string][Math]::Round($_.BackupStats.GetCompressX(), 1)) + "x" } }
            },
            Result  | ConvertTo-HTML -Fragment
            $bodySessSuccRp = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccRp
        } Else {
            $bodySessSuccRp = $successSessionsRp | Sort-Object Creationtime | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
            @{Name = "Start Time"; Expression = { $_.CreationTime } },
            @{Name = "Stop Time"; Expression = { $_.EndTime } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            Result | ConvertTo-HTML -Fragment
            $bodySessSuccRp = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccRp
        }
    }
}

## Gathering tasks after session info has been recorded due to Veeam issue
# Gather all Replication Tasks from Sessions within time frame
$taskListRp = @()
$taskListRp += $sessListRp | Get-VBRTaskSession
$successTasksRp = @($taskListRp | Where-Object { $_.Status -eq "Success" })
$wfTasksRp = @($taskListRp | Where-Object { $_.Status -match "Warning|Failed" })
$runningTasksRp = @()
$runningTasksRp += $runningSessionsRp | Get-VBRTaskSession | Where-Object { $_.Status -match "Pending|InProgress" }

# Get Replication Tasks
$bodyAllTasksRp = $null
If ($showAllTasksRp) {
    If ($taskListRp.count -gt 0) {
        If ($showDetailedRp) {
            $arrAllTasksRp = $taskListRp | Select-Object @{Name = "VM Name"; Expression = { $_.Name } },
            @{Name = "Job Name"; Expression = { $_.JobSess.Name } },
            @{Name = "Start Time"; Expression = { $_.Progress.StartTimeLocal } },
            @{Name = "Stop Time"; Expression = { If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") { "-" } Else { $_.Progress.StopTimeLocal } } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Progress.AvgSpeed / 1MB, 2) } },
            @{Name = "Total (GB)"; Expression = { [Math]::Round($_.Progress.ProcessedSize / 1GB, 2) } },
            @{Name = "Processed (GB)"; Expression = { [Math]::Round($_.Progress.ProcessedUsedSize / 1GB, 2) } },
            @{Name = "Data Read (GB)"; Expression = { [Math]::Round($_.Progress.ReadSize / 1GB, 2) } },
            @{Name = "Transferred (GB)"; Expression = { [Math]::Round($_.Progress.TransferedSize / 1GB, 2) } },
            @{Name = "Details"; Expression = { ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } }, Status
            $bodyAllTasksRp = $arrAllTasksRp | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
            If ($arrAllTasksRp.Status -match "Failed") {
                $allTasksRpHead = $subHead01err
            } ElseIf ($arrAllTasksRp.Status -match "Warning") {
                $allTasksRpHead = $subHead01war
            } ElseIf ($arrAllTasksRp.Status -match "Success") {
                $allTasksRpHead = $subHead01suc
            } Else {
                $allTasksRpHead = $subHead01
            }
            $bodyAllTasksRp = $allTasksRpHead + "Replication Tasks" + $subHead02 + $bodyAllTasksRp
        } Else {
            $arrAllTasksRp = $taskListRp | Select-Object @{Name = "VM Name"; Expression = { $_.Name } },
            @{Name = "Job Name"; Expression = { $_.JobSess.Name } },
            @{Name = "Start Time"; Expression = { $_.Progress.StartTimeLocal } },
            @{Name = "Stop Time"; Expression = { If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") { "-" } Else { $_.Progress.StopTimeLocal } } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Details"; Expression = { ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } }, Status
            $bodyAllTasksRp = $arrAllTasksRp | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
            If ($arrAllTasksRp.Status -match "Failed") {
                $allTasksRpHead = $subHead01err
            } ElseIf ($arrAllTasksRp.Status -match "Warning") {
                $allTasksRpHead = $subHead01war
            } ElseIf ($arrAllTasksRp.Status -match "Success") {
                $allTasksRpHead = $subHead01suc
            } Else {
                $allTasksRpHead = $subHead01
            }
            $bodyAllTasksRp = $allTasksRpHead + "Replication Tasks" + $subHead02 + $bodyAllTasksRp
        }
    }
}

# Get Running Replication Tasks
$bodyTasksRunningRp = $null
If ($showRunningTasksRp) {
    If ($runningTasksRp.count -gt 0) {
        $bodyTasksRunningRp = $runningTasksRp | Select-Object @{Name = "VM Name"; Expression = { $_.Name } },
        @{Name = "Job Name"; Expression = { $_.JobSess.Name } },
        @{Name = "Start Time"; Expression = { $_.Info.Progress.StartTimeLocal } },
        @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
        @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Progress.AvgSpeed / 1MB, 2) } },
        @{Name = "Read (GB)"; Expression = { [Math]::Round($_.Progress.ReadSize / 1GB, 2) } },
        @{Name = "Transferred (GB)"; Expression = { [Math]::Round($_.Progress.TransferedSize / 1GB, 2) } },
        Status | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
        $bodyTasksRunningRp = $subHead01 + "Running Replication Tasks" + $subHead02 + $bodyTasksRunningRp
    }
}

# Get Replication Tasks with Warnings or Failures
$bodyTaskWFRp = $null
If ($showTaskWFRp) {
    If ($wfTasksRp.count -gt 0) {
        If ($showDetailedRp) {
            $arrTaskWFRp = $wfTasksRp | Select-Object @{Name = "VM Name"; Expression = { $_.Name } },
            @{Name = "Job Name"; Expression = { $_.JobSess.Name } },
            @{Name = "Start Time"; Expression = { $_.Progress.StartTimeLocal } },
            @{Name = "Stop Time"; Expression = { $_.Progress.StopTimeLocal } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Progress.AvgSpeed / 1MB, 2) } },
            @{Name = "Total (GB)"; Expression = { [Math]::Round($_.Progress.ProcessedSize / 1GB, 2) } },
            @{Name = "Processed (GB)"; Expression = { [Math]::Round($_.Progress.ProcessedUsedSize / 1GB, 2) } },
            @{Name = "Data Read (GB)"; Expression = { [Math]::Round($_.Progress.ReadSize / 1GB, 2) } },
            @{Name = "Transferred (GB)"; Expression = { [Math]::Round($_.Progress.TransferedSize / 1GB, 2) } },
            @{Name = "Details"; Expression = { ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } }, Status
            $bodyTaskWFRp = $arrTaskWFRp | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
            If ($arrTaskWFRp.Status -match "Failed") {
                $taskWFRpHead = $subHead01err
            } ElseIf ($arrTaskWFRp.Status -match "Warning") {
                $taskWFRpHead = $subHead01war
            } ElseIf ($arrTaskWFRp.Status -match "Success") {
                $taskWFRpHead = $subHead01suc
            } Else {
                $taskWFRpHead = $subHead01
            }
            $bodyTaskWFRp = $taskWFRpHead + "Replication Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFRp
        } Else {
            $arrTaskWFRp = $wfTasksRp | Select-Object @{Name = "VM Name"; Expression = { $_.Name } },
            @{Name = "Job Name"; Expression = { $_.JobSess.Name } },
            @{Name = "Start Time"; Expression = { $_.Progress.StartTimeLocal } },
            @{Name = "Stop Time"; Expression = { $_.Progress.StopTimeLocal } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Details"; Expression = { ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } }, Status
            $bodyTaskWFRp = $arrTaskWFRp | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
            If ($arrTaskWFRp.Status -match "Failed") {
                $taskWFRpHead = $subHead01err
            } ElseIf ($arrTaskWFRp.Status -match "Warning") {
                $taskWFRpHead = $subHead01war
            } ElseIf ($arrTaskWFRp.Status -match "Success") {
                $taskWFRpHead = $subHead01suc
            } Else {
                $taskWFRpHead = $subHead01
            }
            $bodyTaskWFRp = $taskWFRpHead + "Replication Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFRp
        }
    }
}

# Get Successful Replication Tasks
$bodyTaskSuccRp = $null
If ($showTaskSuccessRp) {
    If ($successTasksRp.count -gt 0) {
        If ($showDetailedRp) {
            $bodyTaskSuccRp = $successTasksRp | Select-Object @{Name = "VM Name"; Expression = { $_.Name } },
            @{Name = "Job Name"; Expression = { $_.JobSess.Name } },
            @{Name = "Start Time"; Expression = { $_.Progress.StartTimeLocal } },
            @{Name = "Stop Time"; Expression = { $_.Progress.StopTimeLocal } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Progress.AvgSpeed / 1MB, 2) } },
            @{Name = "Total (GB)"; Expression = { [Math]::Round($_.Progress.ProcessedSize / 1GB, 2) } },
            @{Name = "Processed (GB)"; Expression = { [Math]::Round($_.Progress.ProcessedUsedSize / 1GB, 2) } },
            @{Name = "Data Read (GB)"; Expression = { [Math]::Round($_.Progress.ReadSize / 1GB, 2) } },
            @{Name = "Transferred (GB)"; Expression = { [Math]::Round($_.Progress.TransferedSize / 1GB, 2) } },
            Status | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
            $bodyTaskSuccRp = $subHead01suc + "Successful Replication Tasks" + $subHead02 + $bodyTaskSuccRp
        } Else {
            $bodyTaskSuccRp = $successTasksRp | Select-Object @{Name = "VM Name"; Expression = { $_.Name } },
            @{Name = "Job Name"; Expression = { $_.JobSess.Name } },
            @{Name = "Start Time"; Expression = { $_.Progress.StartTimeLocal } },
            @{Name = "Stop Time"; Expression = { $_.Progress.StopTimeLocal } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            Status | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
            $bodyTaskSuccRp = $subHead01suc + "Successful Replication Tasks" + $subHead02 + $bodyTaskSuccRp
        }
    }
}

# Get Backup Copy Summary Info
$bodySummaryBc = $null
If ($showSummaryBc) {
    $vbrMasterHash = @{
        "Sessions"    = If ($sessListBc) { @($sessListBc).Count } Else { 0 }
        "Read"        = $totalReadBc
        "Transferred" = $totalXferBc
        "Successful"  = @($successSessionsBc).Count
        "Warning"     = @($warningSessionsBc).Count
        "Fails"       = @($failsSessionsBc).Count
        "Working"     = @($workingSessionsBc).Count
        "Idle"        = @($idleSessionsBc).Count
    }
    $vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
    If ($onlyLastBc) {
        $total = "Jobs Run"
    } Else {
        $total = "Total Sessions"
    }
    $arrSummaryBc = $vbrMasterObj | Select-Object @{Name = $total; Expression = { $_.Sessions } },
    @{Name = "Read (GB)"; Expression = { $_.Read } }, @{Name = "Transferred (GB)"; Expression = { $_.Transferred } },
    @{Name = "Idle"; Expression = { $_.Idle } },
    @{Name = "Working"; Expression = { $_.Working } }, @{Name = "Successful"; Expression = { $_.Successful } },
    @{Name = "Warnings"; Expression = { $_.Warning } }, @{Name = "Failures"; Expression = { $_.Fails } }
    $bodySummaryBc = $arrSummaryBc | ConvertTo-HTML -Fragment
    If ($arrSummaryBc.Failures -gt 0) {
        $summaryBcHead = $subHead01err
    } ElseIf ($arrSummaryBc.Warnings -gt 0) {
        $summaryBcHead = $subHead01war
    } ElseIf ($arrSummaryBc.Successful -gt 0) {
        $summaryBcHead = $subHead01suc
    } Else {
        $summaryBcHead = $subHead01
    }
    $bodySummaryBc = $summaryBcHead + "Backup Copy Results Summary" + $subHead02 + $bodySummaryBc
}

# Get Backup Copy Job Status
$bodyJobsBc = $null
If ($showJobsBc) {
    If ($allJobsBc.count -gt 0) {
        $bodyJobsBc = @()
        Foreach ($BcJob in $allJobsBc) {
            $bodyJobsBc += $BcJob | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
            @{Name = "Enabled"; Expression = { $_.Info.IsScheduleEnabled } },
            @{Name = "Type"; Expression = { $_.TypeToString } },
            @{Name = "Status"; Expression = {
                    If ($BcJob.IsRunning) {
                        $currentSess = $BcJob.FindLastSession()
                        If ($currentSess.State -eq "Working") {
                            $csessPercent = $currentSess.Progress.Percents
                            $csessSpeed = [Math]::Round($currentSess.Progress.AvgSpeed / 1MB, 2)
                            $cStatus = "$($csessPercent)% completed at $($csessSpeed) MB/s"
                            $cStatus
                        } Else {
                            $currentSess.State
                        }
                    } Else {
                        "Stopped"
                    }
                }
            },
            @{Name = "Target Repo"; Expression = {
                    If ($($repoList | Where-Object { $_.Id -eq $BcJob.Info.TargetRepositoryId }).Name) { $($repoList | Where-Object { $_.Id -eq $BcJob.Info.TargetRepositoryId }).Name }
                    Else { $($repoListSo | Where-Object { $_.Id -eq $BcJob.Info.TargetRepositoryId }).Name } }
            },
            @{Name = "Next Run"; Expression = {
                    If ($_.IsScheduleEnabled -eq $false) { "<Disabled>" }
                    ElseIf ($_.Options.JobOptions.RunManually) { "<not scheduled>" }
                    ElseIf ($_.ScheduleOptions.IsContinious) { "<Continious>" }
                    ElseIf ($_.ScheduleOptions.OptionsScheduleAfterJob.IsEnabled) { "After [" + $(($allJobs + $allJobsTp) | Where-Object { $_.Id -eq $BcJob.Info.ParentScheduleId }).Name + "]" }
                    Else { $_.ScheduleOptions.NextRun } }
            },
            @{Name = "Last Result"; Expression = { If ($_.Info.LatestStatus -eq "None") { "" }Else { $_.Info.LatestStatus } } }
        }
        $bodyJobsBc = $bodyJobsBc | Sort-Object "Next Run", "Job Name" | ConvertTo-HTML -Fragment
        $bodyJobsBc = $subHead01 + "Backup Copy Job Status" + $subHead02 + $bodyJobsBc
    }
}

# Get Backup Copy Job Size
$bodyJobSizeBc = $null
If ($showBackupSizeBc) {
    If ($backupsBc.count -gt 0) {
        $bodyJobSizeBc = Get-BackupSize -backups $backupsBc | Sort-Object JobName | Select-Object @{Name = "Job Name"; Expression = { $_.JobName } },
        @{Name = "VM Count"; Expression = { $_.VMCount } },
        @{Name = "Repository"; Expression = { $_.Repo } },
        @{Name = "Data Size (GB)"; Expression = { $_.DataSize } },
        @{Name = "Backup Size (GB)"; Expression = { $_.BackupSize } } | ConvertTo-HTML -Fragment
        $bodyJobSizeBc = $subHead01 + "Backup Copy Job Size" + $subHead02 + $bodyJobSizeBc
    }
}

# Get All Backup Copy Sessions
$bodyAllSessBc = $null
If ($showAllSessBc) {
    If ($sessListBc.count -gt 0) {
        If ($showDetailedBc) {
            $arrAllSessBc = $sessListBc | Sort-Object Creationtime | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
            @{Name = "State"; Expression = { $_.State } },
            @{Name = "Start Time"; Expression = { $_.CreationTime } },
            @{Name = "Stop Time"; Expression = { If ($_.EndTime -eq "1/1/1900 12:00:00 AM") { "-" } Else { $_.EndTime } } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Info.Progress.AvgSpeed / 1MB, 2) } },
            @{Name = "Total (GB)"; Expression = { [Math]::Round($_.Info.Progress.ProcessedSize / 1GB, 2) } },
            @{Name = "Processed (GB)"; Expression = { [Math]::Round($_.Info.Progress.ProcessedUsedSize / 1GB, 2) } },
            @{Name = "Data Read (GB)"; Expression = { [Math]::Round($_.Info.Progress.ReadSize / 1GB, 2) } },
            @{Name = "Transferred (GB)"; Expression = { [Math]::Round($_.Info.Progress.TransferedSize / 1GB, 2) } },
            @{Name = "Dedupe"; Expression = {
                    If ($_.Progress.ReadSize -eq 0) { 0 }
                    Else { ([string][Math]::Round($_.BackupStats.GetDedupeX(), 1)) + "x" } }
            },
            @{Name = "Compression"; Expression = {
                    If ($_.Progress.ReadSize -eq 0) { 0 }
                    Else { ([string][Math]::Round($_.BackupStats.GetCompressX(), 1)) + "x" } }
            },
            @{Name = "Details"; Expression = { ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } }, Result
            $bodyAllSessBc = $arrAllSessBc | ConvertTo-HTML -Fragment
            If ($arrAllSessBc.Result -match "Failed") {
                $allSessBcHead = $subHead01err
            } ElseIf ($arrAllSessBc.Result -match "Warning") {
                $allSessBcHead = $subHead01war
            } ElseIf ($arrAllSessBc.Result -match "Success") {
                $allSessBcHead = $subHead01suc
            } Else {
                $allSessBcHead = $subHead01
            }
            $bodyAllSessBc = $allSessBcHead + "Backup Copy Sessions" + $subHead02 + $bodyAllSessBc
        } Else {
            $arrAllSessBc = $sessListBc | Sort-Object Creationtime | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
            @{Name = "State"; Expression = { $_.State } },
            @{Name = "Start Time"; Expression = { $_.CreationTime } },
            @{Name = "Stop Time"; Expression = { If ($_.EndTime -eq "1/1/1900 12:00:00 AM") { "-" } Else { $_.EndTime } } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Details"; Expression = { ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } }, Result
            $bodyAllSessBc = $arrAllSessBc | ConvertTo-HTML -Fragment
            If ($arrAllSessBc.Result -match "Failed") {
                $allSessBcHead = $subHead01err
            } ElseIf ($arrAllSessBc.Result -match "Warning") {
                $allSessBcHead = $subHead01war
            } ElseIf ($arrAllSessBc.Result -match "Success") {
                $allSessBcHead = $subHead01suc
            } Else {
                $allSessBcHead = $subHead01
            }
            $bodyAllSessBc = $allSessBcHead + "Backup Copy Sessions" + $subHead02 + $bodyAllSessBc
        }
    }
}

# Get Idle Backup Copy Sessions
$bodySessIdleBc = $null
If ($showIdleBc) {
    If ($idleSessionsBc.count -gt 0) {
        If ($onlyLastBc) {
            $headerIdle = "Idle Backup Copy Jobs"
        } Else {
            $headerIdle = "Idle Backup Copy Sessions"
        }
        If ($showDetailedBc) {
            $bodySessIdleBc = $idleSessionsBc | Sort-Object Creationtime | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
            @{Name = "Start Time"; Expression = { $_.CreationTime } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date)) } },
            @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Info.Progress.AvgSpeed / 1MB, 2) } },
            @{Name = "Total (GB)"; Expression = { [Math]::Round($_.Info.Progress.ProcessedSize / 1GB, 2) } },
            @{Name = "Processed (GB)"; Expression = { [Math]::Round($_.Info.Progress.ProcessedUsedSize / 1GB, 2) } },
            @{Name = "Data Read (GB)"; Expression = { [Math]::Round($_.Info.Progress.ReadSize / 1GB, 2) } },
            @{Name = "Transferred (GB)"; Expression = { [Math]::Round($_.Info.Progress.TransferedSize / 1GB, 2) } },
            @{Name = "Dedupe"; Expression = {
                    If ($_.Progress.ReadSize -eq 0) { 0 }
                    Else { ([string][Math]::Round($_.BackupStats.GetDedupeX(), 1)) + "x" } }
            },
            @{Name = "Compression"; Expression = {
                    If ($_.Progress.ReadSize -eq 0) { 0 }
                    Else { ([string][Math]::Round($_.BackupStats.GetCompressX(), 1)) + "x" } }
            } | ConvertTo-HTML -Fragment
            $bodySessIdleBc = $subHead01 + $headerIdle + $subHead02 + $bodySessIdleBc
        } Else {
            $bodySessIdleBc = $idleSessionsBc | Sort-Object Creationtime | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
            @{Name = "Start Time"; Expression = { $_.CreationTime } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date)) } } | ConvertTo-HTML -Fragment
            $bodySessIdleBc = $subHead01 + $headerIdle + $subHead02 + $bodySessIdleBc
        }
    }
}

# Get Working Backup Copy Jobs
$bodyRunningBc = $null
If ($showRunningBc) {
    If ($workingSessionsBc.count -gt 0) {
        $bodyRunningBc = $workingSessionsBc | Sort-Object Creationtime | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
        @{Name = "Start Time"; Expression = { $_.CreationTime } },
        @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $(New-TimeSpan $_.Progress.StartTimeLocal $(Get-Date)) } },
        @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Progress.AvgSpeed / 1MB, 2) } },
        @{Name = "Read (GB)"; Expression = { [Math]::Round([Decimal]$_.Progress.ReadSize / 1GB, 2) } },
        @{Name = "Transferred (GB)"; Expression = { [Math]::Round([Decimal]$_.Progress.TransferedSize / 1GB, 2) } },
        @{Name = "% Complete"; Expression = { $_.Progress.Percents } } | ConvertTo-HTML -Fragment
        $bodyRunningBc = $subHead01 + "Working Backup Copy Sessions" + $subHead02 + $bodyRunningBc
    }
}

# Get Backup Copy Sessions with Warnings or Failures
$bodySessWFBc = $null
If ($showWarnFailBc) {
    $sessWF = @($warningSessionsBc + $failsSessionsBc)
    If ($sessWF.count -gt 0) {
        If ($onlyLastBc) {
            $headerWF = "Backup Copy Jobs with Warnings or Failures"
        } Else {
            $headerWF = "Backup Copy Sessions with Warnings or Failures"
        }
        If ($showDetailedBc) {
            $arrSessWFBc = $sessWF | Sort-Object Creationtime | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
            @{Name = "Start Time"; Expression = { $_.CreationTime } },
            @{Name = "Stop Time"; Expression = { $_.EndTime } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Info.Progress.AvgSpeed / 1MB, 2) } },
            @{Name = "Total (GB)"; Expression = { [Math]::Round($_.Info.Progress.ProcessedSize / 1GB, 2) } },
            @{Name = "Processed (GB)"; Expression = { [Math]::Round($_.Info.Progress.ProcessedUsedSize / 1GB, 2) } },
            @{Name = "Data Read (GB)"; Expression = { [Math]::Round($_.Info.Progress.ReadSize / 1GB, 2) } },
            @{Name = "Transferred (GB)"; Expression = { [Math]::Round($_.Info.Progress.TransferedSize / 1GB, 2) } },
            @{Name = "Dedupe"; Expression = {
                    If ($_.Progress.ReadSize -eq 0) { 0 }
                    Else { ([string][Math]::Round($_.BackupStats.GetDedupeX(), 1)) + "x" } }
            },
            @{Name = "Compression"; Expression = {
                    If ($_.Progress.ReadSize -eq 0) { 0 }
                    Else { ([string][Math]::Round($_.BackupStats.GetCompressX(), 1)) + "x" } }
            },
            @{Name = "Details"; Expression = {
                    If ($_.GetDetails() -eq "") { $_ | Get-VBRTaskSession | ForEach-Object { If ($_.GetDetails()) { $_.Name + ": " + ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } } }
                    Else { ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } }
            }, Result
            $bodySessWFBc = $arrSessWFBc | ConvertTo-HTML -Fragment
            If ($arrSessWFBc.Result -match "Failed") {
                $sessWFBcHead = $subHead01err
            } ElseIf ($arrSessWFBc.Result -match "Warning") {
                $sessWFBcHead = $subHead01war
            } ElseIf ($arrSessWFBc.Result -match "Success") {
                $sessWFBcHead = $subHead01suc
            } Else {
                $sessWFBcHead = $subHead01
            }
            $bodySessWFBc = $sessWFBcHead + $headerWF + $subHead02 + $bodySessWFBc
        } Else {
            $arrSessWFBc = $sessWF | Sort-Object Creationtime | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
            @{Name = "Start Time"; Expression = { $_.CreationTime } },
            @{Name = "Stop Time"; Expression = { $_.EndTime } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Details"; Expression = {
                    If ($_.GetDetails() -eq "") { $_ | Get-VBRTaskSession | ForEach-Object { If ($_.GetDetails()) { $_.Name + ": " + ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } } }
                    Else { ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } }
            }, Result
            $bodySessWFBc = $arrSessWFBc | ConvertTo-HTML -Fragment
            If ($arrSessWFBc.Result -match "Failed") {
                $sessWFBcHead = $subHead01err
            } ElseIf ($arrSessWFBc.Result -match "Warning") {
                $sessWFBcHead = $subHead01war
            } ElseIf ($arrSessWFBc.Result -match "Success") {
                $sessWFBcHead = $subHead01suc
            } Else {
                $sessWFBcHead = $subHead01
            }
            $bodySessWFBc = $sessWFBcHead + $headerWF + $subHead02 + $bodySessWFBc
        }
    }
}

# Get Successful Backup Copy Sessions
$bodySessSuccBc = $null
If ($showSuccessBc) {
    If ($successSessionsBc.count -gt 0) {
        If ($onlyLastBc) {
            $headerSucc = "Successful Backup Copy Jobs"
        } Else {
            $headerSucc = "Successful Backup Copy Sessions"
        }
        If ($showDetailedBc) {
            $bodySessSuccBc = $successSessionsBc | Sort-Object Creationtime | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
            @{Name = "Start Time"; Expression = { $_.CreationTime } },
            @{Name = "Stop Time"; Expression = { $_.EndTime } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Info.Progress.AvgSpeed / 1MB, 2) } },
            @{Name = "Total (GB)"; Expression = { [Math]::Round($_.Info.Progress.ProcessedSize / 1GB, 2) } },
            @{Name = "Processed (GB)"; Expression = { [Math]::Round($_.Info.Progress.ProcessedUsedSize / 1GB, 2) } },
            @{Name = "Data Read (GB)"; Expression = { [Math]::Round($_.Info.Progress.ReadSize / 1GB, 2) } },
            @{Name = "Transferred (GB)"; Expression = { [Math]::Round($_.Info.Progress.TransferedSize / 1GB, 2) } },
            @{Name = "Dedupe"; Expression = {
                    If ($_.Progress.ReadSize -eq 0) { 0 }
                    Else { ([string][Math]::Round($_.BackupStats.GetDedupeX(), 1)) + "x" } }
            },
            @{Name = "Compression"; Expression = {
                    If ($_.Progress.ReadSize -eq 0) { 0 }
                    Else { ([string][Math]::Round($_.BackupStats.GetCompressX(), 1)) + "x" } }
            },
            Result  | ConvertTo-HTML -Fragment
            $bodySessSuccBc = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccBc
        } Else {
            $bodySessSuccBc = $successSessionsBc | Sort-Object Creationtime | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
            @{Name = "Start Time"; Expression = { $_.CreationTime } },
            @{Name = "Stop Time"; Expression = { $_.EndTime } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            Result | ConvertTo-HTML -Fragment
            $bodySessSuccBc = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccBc
        }
    }
}

## Gathering tasks after session info has been recorded due to Veeam issue
# Gather all Backup Copy Tasks from Sessions within time frame
$taskListBc = @()
$taskListBc += $sessListBc | Get-VBRTaskSession
$successTasksBc = @($taskListBc | Where-Object { $_.Status -eq "Success" })
$wfTasksBc = @($taskListBc | Where-Object { $_.Status -match "Warning|Failed" })
$pendingTasksBc = @($taskListBc | Where-Object { $_.Status -eq "Pending" })
$runningTasksBc = @($taskListBc | Where-Object { $_.Status -eq "InProgress" })

# Get All Backup Copy Tasks
$bodyAllTasksBc = $null
If ($showAllTasksBc) {
    If ($taskListBc.count -gt 0) {
        If ($showDetailedBc) {
            $arrAllTasksBc = $taskListBc | Select-Object @{Name = "VM Name"; Expression = { $_.Name } },
            @{Name = "Job Name"; Expression = { $_.JobSess.Name } },
            @{Name = "Start Time"; Expression = { $_.Progress.StartTimeLocal } },
            @{Name = "Stop Time"; Expression = { If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") { "-" } Else { $_.Progress.StopTimeLocal } } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Progress.AvgSpeed / 1MB, 2) } },
            @{Name = "Total (GB)"; Expression = { [Math]::Round($_.Progress.ProcessedSize / 1GB, 2) } },
            @{Name = "Processed (GB)"; Expression = { [Math]::Round($_.Progress.ProcessedUsedSize / 1GB, 2) } },
            @{Name = "Data Read (GB)"; Expression = { [Math]::Round($_.Progress.ReadSize / 1GB, 2) } },
            @{Name = "Transferred (GB)"; Expression = { [Math]::Round($_.Progress.TransferedSize / 1GB, 2) } },
            @{Name = "Details"; Expression = { ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } }, Status
            $bodyAllTasksBc = $arrAllTasksBc | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
            If ($arrAllTasksBc.Status -match "Failed") {
                $allTasksBcHead = $subHead01err
            } ElseIf ($arrAllTasksBc.Status -match "Warning") {
                $allTasksBcHead = $subHead01war
            } ElseIf ($arrAllTasksBc.Status -match "Success") {
                $allTasksBcHead = $subHead01suc
            } Else {
                $allTasksBcHead = $subHead01
            }
            $bodyAllTasksBc = $allTasksBcHead + "Backup Copy Tasks" + $subHead02 + $bodyAllTasksBc
        } Else {
            $arrAllTasksBc = $taskListBc | Select-Object @{Name = "VM Name"; Expression = { $_.Name } },
            @{Name = "Job Name"; Expression = { $_.JobSess.Name } },
            @{Name = "Start Time"; Expression = { $_.Progress.StartTimeLocal } },
            @{Name = "Stop Time"; Expression = { If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") { "-" } Else { $_.Progress.StopTimeLocal } } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Details"; Expression = { ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } }, Status
            $bodyAllTasksBc = $arrAllTasksBc | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
            If ($arrAllTasksBc.Status -match "Failed") {
                $allTasksBcHead = $subHead01err
            } ElseIf ($arrAllTasksBc.Status -match "Warning") {
                $allTasksBcHead = $subHead01war
            } ElseIf ($arrAllTasksBc.Status -match "Success") {
                $allTasksBcHead = $subHead01suc
            } Else {
                $allTasksBcHead = $subHead01
            }
            $bodyAllTasksBc = $allTasksBcHead + "Backup Copy Tasks" + $subHead02 + $bodyAllTasksBc
        }
    }
}

# Get Pending Backup Copy Tasks
$bodyTasksPendingBc = $null
If ($showPendingTasksBc) {
    If ($pendingTasksBc.count -gt 0) {
        $bodyTasksPendingBc = $pendingTasksBc | Select-Object @{Name = "VM Name"; Expression = { $_.Name } },
        @{Name = "Job Name"; Expression = { $_.JobSess.Name } },
        @{Name = "Start Time"; Expression = { $_.Info.Progress.StartTimeLocal } },
        @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
        @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Progress.AvgSpeed / 1MB, 2) } },
        @{Name = "Read (GB)"; Expression = { [Math]::Round($_.Progress.ReadSize / 1GB, 2) } },
        @{Name = "Transferred (GB)"; Expression = { [Math]::Round($_.Progress.TransferedSize / 1GB, 2) } },
        Status | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
        $bodyTasksPendingBc = $subHead01 + "Pending Backup Copy Tasks" + $subHead02 + $bodyTasksPendingBc
    }
}

# Get Working Backup Copy Tasks
$bodyTasksRunningBc = $null
If ($showRunningTasksBc) {
    If ($runningTasksBc.count -gt 0) {
        $bodyTasksRunningBc = $runningTasksBc | Select-Object @{Name = "VM Name"; Expression = { $_.Name } },
        @{Name = "Job Name"; Expression = { $_.JobSess.Name } },
        @{Name = "Start Time"; Expression = { $_.Info.Progress.StartTimeLocal } },
        @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
        @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Progress.AvgSpeed / 1MB, 2) } },
        @{Name = "Read (GB)"; Expression = { [Math]::Round($_.Progress.ReadSize / 1GB, 2) } },
        @{Name = "Transferred (GB)"; Expression = { [Math]::Round($_.Progress.TransferedSize / 1GB, 2) } },
        Status | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
        $bodyTasksRunningBc = $subHead01 + "Working Backup Copy Tasks" + $subHead02 + $bodyTasksRunningBc
    }
}

# Get Backup Copy Tasks with Warnings or Failures
$bodyTaskWFBc = $null
If ($showTaskWFBc) {
    If ($wfTasksBc.count -gt 0) {
        If ($showDetailedBc) {
            $arrTaskWFBc = $wfTasksBc | Select-Object @{Name = "VM Name"; Expression = { $_.Name } },
            @{Name = "Job Name"; Expression = { $_.JobSess.Name } },
            @{Name = "Start Time"; Expression = { $_.Progress.StartTimeLocal } },
            @{Name = "Stop Time"; Expression = { $_.Progress.StopTimeLocal } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Progress.AvgSpeed / 1MB, 2) } },
            @{Name = "Total (GB)"; Expression = { [Math]::Round($_.Progress.ProcessedSize / 1GB, 2) } },
            @{Name = "Processed (GB)"; Expression = { [Math]::Round($_.Progress.ProcessedUsedSize / 1GB, 2) } },
            @{Name = "Data Read (GB)"; Expression = { [Math]::Round($_.Progress.ReadSize / 1GB, 2) } },
            @{Name = "Transferred (GB)"; Expression = { [Math]::Round($_.Progress.TransferedSize / 1GB, 2) } },
            @{Name = "Details"; Expression = { ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } }, Status
            $bodyTaskWFBc = $arrTaskWFBc | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
            If ($arrTaskWFBc.Status -match "Failed") {
                $taskWFBcHead = $subHead01err
            } ElseIf ($arrTaskWFBc.Status -match "Warning") {
                $taskWFBcHead = $subHead01war
            } ElseIf ($arrTaskWFBc.Status -match "Success") {
                $taskWFBcHead = $subHead01suc
            } Else {
                $taskWFBcHead = $subHead01
            }
            $bodyTaskWFBc = $taskWFBcHead + "Backup Copy Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFBc
        } Else {
            $arrTaskWFBc = $wfTasksBc | Select-Object @{Name = "VM Name"; Expression = { $_.Name } },
            @{Name = "Job Name"; Expression = { $_.JobSess.Name } },
            @{Name = "Start Time"; Expression = { $_.Progress.StartTimeLocal } },
            @{Name = "Stop Time"; Expression = { $_.Progress.StopTimeLocal } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Details"; Expression = { ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } }, Status
            $bodyTaskWFBc = $arrTaskWFBc | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
            If ($arrTaskWFBc.Status -match "Failed") {
                $taskWFBcHead = $subHead01err
            } ElseIf ($arrTaskWFBc.Status -match "Warning") {
                $taskWFBcHead = $subHead01war
            } ElseIf ($arrTaskWFBc.Status -match "Success") {
                $taskWFBcHead = $subHead01suc
            } Else {
                $taskWFBcHead = $subHead01
            }
            $bodyTaskWFBc = $taskWFBcHead + "Backup Copy Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFBc
        }
    }
}

# Get Successful Backup Copy Tasks
$bodyTaskSuccBc = $null
If ($showTaskSuccessBc) {
    If ($successTasksBc.count -gt 0) {
        If ($showDetailedBc) {
            $bodyTaskSuccBc = $successTasksBc | Select-Object @{Name = "VM Name"; Expression = { $_.Name } },
            @{Name = "Job Name"; Expression = { $_.JobSess.Name } },
            @{Name = "Start Time"; Expression = { $_.Progress.StartTimeLocal } },
            @{Name = "Stop Time"; Expression = {
                    If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") { "-" }
                    Else { $_.Progress.StopTimeLocal }
                }
            },
            @{Name = "Duration (HH:MM:SS)"; Expression = {
                    If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") { "-" }
                    Else { Get-Duration -ts $_.Progress.Duration }
                }
            },
            @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Progress.AvgSpeed / 1MB, 2) } },
            @{Name = "Total (GB)"; Expression = { [Math]::Round($_.Progress.ProcessedSize / 1GB, 2) } },
            @{Name = "Processed (GB)"; Expression = { [Math]::Round($_.Progress.ProcessedUsedSize / 1GB, 2) } },
            @{Name = "Data Read (GB)"; Expression = { [Math]::Round($_.Progress.ReadSize / 1GB, 2) } },
            @{Name = "Transferred (GB)"; Expression = { [Math]::Round($_.Progress.TransferedSize / 1GB, 2) } },
            Status | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
            $bodyTaskSuccBc = $subHead01suc + "Successful Backup Copy Tasks" + $subHead02 + $bodyTaskSuccBc
        } Else {
            $bodyTaskSuccBc = $successTasksBc | Select-Object @{Name = "VM Name"; Expression = { $_.Name } },
            @{Name = "Job Name"; Expression = { $_.JobSess.Name } },
            @{Name = "Start Time"; Expression = { $_.Progress.StartTimeLocal } },
            @{Name = "Stop Time"; Expression = {
                    If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") { "-" }
                    Else { $_.Progress.StopTimeLocal }
                }
            },
            @{Name = "Duration (HH:MM:SS)"; Expression = {
                    If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") { "-" }
                    Else { Get-Duration -ts $_.Progress.Duration }
                }
            },
            Status | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
            $bodyTaskSuccBc = $subHead01suc + "Successful Backup Copy Tasks" + $subHead02 + $bodyTaskSuccBc
        }
    }
}

# Get Tape Backup Summary Info
$bodySummaryTp = $null
If ($showSummaryTp) {
    $vbrMasterHash = @{
        "Sessions"    = If ($sessListTp) { @($sessListTp).Count } Else { 0 }
        "Read"        = $totalReadTp
        "Transferred" = $totalXferTp
        "Successful"  = @($successSessionsTp).Count
        "Warning"     = @($warningSessionsTp).Count
        "Fails"       = @($failsSessionsTp).Count
        "Working"     = @($workingSessionsTp).Count
        "Idle"        = @($idleSessionsTp).Count
        "Waiting"     = @($waitingSessionsTp).Count
    }
    $vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
    If ($onlyLastTp) {
        $total = "Jobs Run"
    } Else {
        $total = "Total Sessions"
    }
    $arrSummaryTp = $vbrMasterObj | Select-Object @{Name = $total; Expression = { $_.Sessions } },
    @{Name = "Read (GB)"; Expression = { $_.Read } }, @{Name = "Transferred (GB)"; Expression = { $_.Transferred } },
    @{Name = "Idle"; Expression = { $_.Idle } }, @{Name = "Waiting"; Expression = { $_.Waiting } },
    @{Name = "Working"; Expression = { $_.Working } }, @{Name = "Successful"; Expression = { $_.Successful } },
    @{Name = "Warnings"; Expression = { $_.Warning } }, @{Name = "Failures"; Expression = { $_.Fails } }
    $bodySummaryTp = $arrSummaryTp | ConvertTo-HTML -Fragment
    If ($arrSummaryTp.Failures -gt 0) {
        $summaryTpHead = $subHead01err
    } ElseIf ($arrSummaryTp.Warnings -gt 0 -or $arrSummaryTp.Waiting -gt 0) {
        $summaryTpHead = $subHead01war
    } ElseIf ($arrSummaryTp.Successful -gt 0) {
        $summaryTpHead = $subHead01suc
    } Else {
        $summaryTpHead = $subHead01
    }
    $bodySummaryTp = $summaryTpHead + "Tape Backup Results Summary" + $subHead02 + $bodySummaryTp
}

# Get Tape Backup Job Status
$bodyJobsTp = $null
If ($showJobsTp) {
    If ($allJobsTp.count -gt 0) {
        $bodyJobsTp = @()
        Foreach ($tpJob in $allJobsTp) {
            $bodyJobsTp += $tpJob | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
            @{Name = "Job Type"; Expression = { $_.Type } }, @{Name = "Media Pool"; Expression = { $_.Target } },
            @{Name = "Status"; Expression = { $_.LastState } },
            @{Name = "Next Run"; Expression = {
                    If ($_.ScheduleOptions.Type -eq "AfterNewBackup") { "<Continious>" }
                    ElseIf ($_.ScheduleOptions.Type -eq "AfterJob") { "After [" + $(($allJobs + $allJobsTp) | Where-Object { $_.Id -eq $tpJob.ScheduleOptions.JobId }).Name + "]" }
                    ElseIf ($_.NextRun) { $_.NextRun }
                    Else { "<not scheduled>" } }
            },
            @{Name = "Last Result"; Expression = { If ($_.LastResult -eq "None") { "" }Else { $_.LastResult } } }
        }
        $bodyJobsTp = $bodyJobsTp | Sort-Object "Next Run", "Job Name" | ConvertTo-HTML -Fragment
        $bodyJobsTp = $subHead01 + "Tape Backup Job Status" + $subHead02 + $bodyJobsTp
    }
}

# Get Tape Backup Sessions
$bodyAllSessTp = $null
If ($showAllSessTp) {
    If ($sessListTp.count -gt 0) {
        If ($showDetailedTp) {
            $arrAllSessTp = $sessListTp | Sort-Object Creationtime | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
            @{Name = "State"; Expression = { $_.State } },
            @{Name = "Start Time"; Expression = { $_.CreationTime } },
            @{Name = "Stop Time"; Expression = { If ($_.EndTime -eq "1/1/1900 12:00:00 AM") { "-" } Else { $_.EndTime } } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Info.Progress.AvgSpeed / 1MB, 2) } },
            @{Name = "Total (GB)"; Expression = { [Math]::Round($_.Info.Progress.ProcessedSize / 1GB, 2) } },
            @{Name = "Data Read (GB)"; Expression = { [Math]::Round($_.Info.Progress.ReadSize / 1GB, 2) } },
            @{Name = "Transferred (GB)"; Expression = { [Math]::Round($_.Info.Progress.TransferedSize / 1GB, 2) } },
            @{Name = "Details"; Expression = { ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } }, Result
            $bodyAllSessTp = $arrAllSessTp | ConvertTo-HTML -Fragment
            If ($arrAllSessTp.Result -match "Failed") {
                $allSessTpHead = $subHead01err
            } ElseIf ($arrAllSessTp.Result -match "Warning" -or $arrAllSessTp.State -match "WaitingTape") {
                $allSessTpHead = $subHead01war
            } ElseIf ($arrAllSessTp.Result -match "Success") {
                $allSessTpHead = $subHead01suc
            } Else {
                $allSessTpHead = $subHead01
            }
            $bodyAllSessTp = $allSessTpHead + "Tape Backup Sessions" + $subHead02 + $bodyAllSessTp
        } Else {
            $arrAllSessTp = $sessListTp | Sort-Object Creationtime | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
            @{Name = "State"; Expression = { $_.State } },
            @{Name = "Start Time"; Expression = { $_.CreationTime } },
            @{Name = "Stop Time"; Expression = { If ($_.EndTime -eq "1/1/1900 12:00:00 AM") { "-" } Else { $_.EndTime } } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Details"; Expression = { ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } }, Result
            $bodyAllSessTp = $arrAllSessTp | ConvertTo-HTML -Fragment
            If ($arrAllSessTp.Result -match "Failed") {
                $allSessTpHead = $subHead01err
            } ElseIf ($arrAllSessTp.Result -match "Warning" -or $arrAllSessTp.State -match "WaitingTape") {
                $allSessTpHead = $subHead01war
            } ElseIf ($arrAllSessTp.Result -match "Success") {
                $allSessTpHead = $subHead01suc
            } Else {
                $allSessTpHead = $subHead01
            }
            $bodyAllSessTp = $allSessTpHead + "Tape Backup Sessions" + $subHead02 + $bodyAllSessTp
        }

        # Due to issue with getting details on tape sessions, we may need to get session info again :-(
        If (($showWaitingTp -or $showIdleTp -or $showRunningTp -or $showWarnFailTp -or $showSuccessTp) -and $showDetailedTp) {
            # Get all Tape Backup Sessions
            $allSessTp = @()
            Foreach ($tpJob in $allJobsTp) {
                $tpSessions = [veeam.backup.core.cbackupsession]::GetByJob($tpJob.id)
                $allSessTp += $tpSessions
            }
            # Gather all Tape Backup Sessions within timeframe
            $sessListTp = @($allSessTp | Where-Object { $_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -match "Working|Idle" })
            If ($null -ne $tapeJob -and $tapeJob -ne "") {
                $allJobsTpTmp = @()
                $sessListTpTmp = @()
                Foreach ($tpJob in $tapeJob) {
                    $allJobsTpTmp += $allJobsTp | Where-Object { $_.Name -like $tpJob }
                    $sessListTpTmp += $sessListTp | Where-Object { $_.JobName -like $tpJob }
                }
                $allJobsTp = $allJobsTpTmp | Sort-Object Id -Unique
                $sessListTp = $sessListTpTmp | Sort-Object Id -Unique
            }
            If ($onlyLastTp) {
                $tempSessListTp = $sessListTp
                $sessListTp = @()
                Foreach ($job in $allJobsTp) {
                    $sessListTp += $tempSessListTp | Where-Object { $_.Jobname -eq $job.name } | Sort-Object EndTime -Descending | Select-Object -First 1
                }
            }
            # Get Tape Backup Session information
            $idleSessionsTp = @($sessListTp | Where-Object { $_.State -eq "Idle" })
            $successSessionsTp = @($sessListTp | Where-Object { $_.Result -eq "Success" })
            $warningSessionsTp = @($sessListTp | Where-Object { $_.Result -eq "Warning" })
            $failsSessionsTp = @($sessListTp | Where-Object { $_.Result -eq "Failed" })
            $workingSessionsTp = @($sessListTp | Where-Object { $_.State -eq "Working" })
            $waitingSessionsTp = @($sessListTp | Where-Object { $_.State -eq "WaitingTape" })
        }
    }
}

# Get Waiting Tape Backup Jobs
$bodyWaitingTp = $null
If ($showWaitingTp) {
    If ($waitingSessionsTp.count -gt 0) {
        $bodyWaitingTp = $waitingSessionsTp | Sort-Object Creationtime | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
        @{Name = "Start Time"; Expression = { $_.CreationTime } },
        @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $(New-TimeSpan $_.Progress.StartTimeLocal $(Get-Date)) } },
        @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Progress.AvgSpeed / 1MB, 2) } },
        @{Name = "Read (GB)"; Expression = { [Math]::Round([Decimal]$_.Progress.ReadSize / 1GB, 2) } },
        @{Name = "Transferred (GB)"; Expression = { [Math]::Round([Decimal]$_.Progress.TransferedSize / 1GB, 2) } },
        @{Name = "% Complete"; Expression = { $_.Progress.Percents } } | ConvertTo-HTML -Fragment
        $bodyWaitingTp = $subHead01war + "Waiting Tape Backup Sessions" + $subHead02 + $bodyWaitingTp
    }
}

# Get Idle Tape Backup Sessions
$bodySessIdleTp = $null
If ($showIdleTp) {
    If ($idleSessionsTp.count -gt 0) {
        If ($onlyLastTp) {
            $headerIdle = "Idle Tape Backup Jobs"
        } Else {
            $headerIdle = "Idle Tape Backup Sessions"
        }
        If ($showDetailedTp) {
            $bodySessIdleTp = $idleSessionsTp | Sort-Object Creationtime | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
            @{Name = "Start Time"; Expression = { $_.CreationTime } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date)) } },
            @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Info.Progress.AvgSpeed / 1MB, 2) } },
            @{Name = "Total (GB)"; Expression = { [Math]::Round($_.Info.Progress.ProcessedSize / 1GB, 2) } },
            @{Name = "Data Read (GB)"; Expression = { [Math]::Round($_.Info.Progress.ReadSize / 1GB, 2) } },
            @{Name = "Transferred (GB)"; Expression = { [Math]::Round($_.Info.Progress.TransferedSize / 1GB, 2) } } | ConvertTo-HTML -Fragment
            $bodySessIdleTp = $subHead01 + $headerIdle + $subHead02 + $bodySessIdleTp
        } Else {
            $bodySessIdleTp = $idleSessionsTp | Sort-Object Creationtime | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
            @{Name = "Start Time"; Expression = { $_.CreationTime } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date)) } } | ConvertTo-HTML -Fragment
            $bodySessIdleTp = $subHead01 + $headerIdle + $subHead02 + $bodySessIdleTp
        }
    }
}

# Get Working Tape Backup Jobs
$bodyRunningTp = $null
If ($showRunningTp) {
    If ($workingSessionsTp.count -gt 0) {
        $bodyRunningTp = $workingSessionsTp | Sort-Object Creationtime | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
        @{Name = "Start Time"; Expression = { $_.CreationTime } },
        @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $(New-TimeSpan $_.Progress.StartTimeLocal $(Get-Date)) } },
        @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Progress.AvgSpeed / 1MB, 2) } },
        @{Name = "Read (GB)"; Expression = { [Math]::Round([Decimal]$_.Progress.ReadSize / 1GB, 2) } },
        @{Name = "Transferred (GB)"; Expression = { [Math]::Round([Decimal]$_.Progress.TransferedSize / 1GB, 2) } },
        @{Name = "% Complete"; Expression = { $_.Progress.Percents } } | ConvertTo-HTML -Fragment
        $bodyRunningTp = $subHead01 + "Working Tape Backup Sessions" + $subHead02 + $bodyRunningTp
    }
}

# Get Tape Backup Sessions with Warnings or Failures
$bodySessWFTp = $null
If ($showWarnFailTp) {
    $sessWF = @($warningSessionsTp + $failsSessionsTp)
    If ($sessWF.count -gt 0) {
        If ($onlyLastTp) {
            $headerWF = "Tape Backup Jobs with Warnings or Failures"
        } Else {
            $headerWF = "Tape Backup Sessions with Warnings or Failures"
        }
        If ($showDetailedTp) {
            $arrSessWFTp = $sessWF | Sort-Object Creationtime | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
            @{Name = "Start Time"; Expression = { $_.CreationTime } },
            @{Name = "Stop Time"; Expression = { $_.EndTime } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Info.Progress.AvgSpeed / 1MB, 2) } },
            @{Name = "Total (GB)"; Expression = { [Math]::Round($_.Info.Progress.ProcessedSize / 1GB, 2) } },
            @{Name = "Data Read (GB)"; Expression = { [Math]::Round($_.Info.Progress.ReadSize / 1GB, 2) } },
            @{Name = "Transferred (GB)"; Expression = { [Math]::Round($_.Info.Progress.TransferedSize / 1GB, 2) } },
            @{Name = "Details"; Expression = {
                    If ($_.GetDetails() -eq "") { $_ | Get-VBRTaskSession | ForEach-Object { If ($_.GetDetails()) { $_.Name + ": " + ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } } }
                    Else { ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } }
            }, Result
            $bodySessWFTp = $arrSessWFTp | ConvertTo-HTML -Fragment
            If ($arrSessWFTp.Result -match "Failed") {
                $sessWFTpHead = $subHead01err
            } ElseIf ($arrSessWFTp.Result -match "Warning") {
                $sessWFTpHead = $subHead01war
            } ElseIf ($arrSessWFTp.Result -match "Success") {
                $sessWFTpHead = $subHead01suc
            } Else {
                $sessWFTpHead = $subHead01
            }
            $bodySessWFTp = $sessWFTpHead + $headerWF + $subHead02 + $bodySessWFTp
        } Else {
            $arrSessWFTp = $sessWF | Sort-Object Creationtime | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
            @{Name = "Start Time"; Expression = { $_.CreationTime } },
            @{Name = "Stop Time"; Expression = { $_.EndTime } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Details"; Expression = {
                    If ($_.GetDetails() -eq "") { $_ | Get-VBRTaskSession | ForEach-Object { If ($_.GetDetails()) { $_.Name + ": " + ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } } }
                    Else { ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } }
            }, Result
            $bodySessWFTp = $arrSessWFTp | ConvertTo-HTML -Fragment
            If ($arrSessWFTp.Result -match "Failed") {
                $sessWFTpHead = $subHead01err
            } ElseIf ($arrSessWFTp.Result -match "Warning") {
                $sessWFTpHead = $subHead01war
            } ElseIf ($arrSessWFTp.Result -match "Success") {
                $sessWFTpHead = $subHead01suc
            } Else {
                $sessWFTpHead = $subHead01
            }
            $bodySessWFTp = $sessWFTpHead + $headerWF + $subHead02 + $bodySessWFTp
        }
    }
}

# Get Successful Tape Backup Sessions
$bodySessSuccTp = $null
If ($showSuccessTp) {
    If ($successSessionsTp.count -gt 0) {
        If ($onlyLastTp) {
            $headerSucc = "Successful Tape Backup Jobs"
        } Else {
            $headerSucc = "Successful Tape Backup Sessions"
        }
        If ($showDetailedTp) {
            $bodySessSuccTp = $successSessionsTp | Sort-Object Creationtime | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
            @{Name = "Start Time"; Expression = { $_.CreationTime } },
            @{Name = "Stop Time"; Expression = { $_.EndTime } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Info.Progress.AvgSpeed / 1MB, 2) } },
            @{Name = "Total (GB)"; Expression = { [Math]::Round($_.Info.Progress.ProcessedSize / 1GB, 2) } },
            @{Name = "Data Read (GB)"; Expression = { [Math]::Round($_.Info.Progress.ReadSize / 1GB, 2) } },
            @{Name = "Transferred (GB)"; Expression = { [Math]::Round($_.Info.Progress.TransferedSize / 1GB, 2) } },
            @{Name = "Details"; Expression = {
                    If ($_.GetDetails() -eq "") { $_ | Get-VBRTaskSession | ForEach-Object { If ($_.GetDetails()) { $_.Name + ": " + ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } } }
                    Else { ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } }
            },
            Result  | ConvertTo-HTML -Fragment
            $bodySessSuccTp = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccTp
        } Else {
            $bodySessSuccTp = $successSessionsTp | Sort-Object Creationtime | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
            @{Name = "Start Time"; Expression = { $_.CreationTime } },
            @{Name = "Stop Time"; Expression = { $_.EndTime } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Details"; Expression = {
                    If ($_.GetDetails() -eq "") { $_ | Get-VBRTaskSession | ForEach-Object { If ($_.GetDetails()) { $_.Name + ": " + ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } } }
                    Else { ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } }
            },
            Result | ConvertTo-HTML -Fragment
            $bodySessSuccTp = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccTp
        }
    }
}

## Gathering tasks after session info has been recorded due to Veeam issue
# Gather all Tape Backup Tasks from Sessions within time frame
$taskListTp = @()
$taskListTp += $sessListTp | Get-VBRTaskSession
$successTasksTp = @($taskListTp | Where-Object { $_.Status -eq "Success" })
$wfTasksTp = @($taskListTp | Where-Object { $_.Status -match "Warning|Failed" })
$pendingTasksTp = @($taskListTp | Where-Object { $_.Status -eq "Pending" })
$runningTasksTp = @($taskListTp | Where-Object { $_.Status -eq "InProgress" })

# Get Tape Backup Tasks
$bodyAllTasksTp = $null
If ($showAllTasksTp) {
    If ($taskListTp.count -gt 0) {
        If ($showDetailedTp) {
            $arrAllTasksTp = $taskListTp | Select-Object @{Name = "Name"; Expression = { $_.Name } },
            @{Name = "Job Name"; Expression = { $_.JobSess.Name } },
            @{Name = "Start Time"; Expression = { $_.Progress.StartTimeLocal } },
            @{Name = "Stop Time"; Expression = { If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") { "-" } Else { $_.Progress.StopTimeLocal } } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Progress.AvgSpeed / 1MB, 2) } },
            @{Name = "Total (GB)"; Expression = { [Math]::Round($_.Progress.ProcessedSize / 1GB, 2) } },
            @{Name = "Data Read (GB)"; Expression = { [Math]::Round($_.Progress.ReadSize / 1GB, 2) } },
            @{Name = "Transferred (GB)"; Expression = { [Math]::Round($_.Progress.TransferedSize / 1GB, 2) } },
            @{Name = "Details"; Expression = { ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } }, Status
            $bodyAllTasksTp = $arrAllTasksTp | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
            If ($arrAllTasksTp.Status -match "Failed") {
                $allTasksTpHead = $subHead01err
            } ElseIf ($arrAllTasksTp.Status -match "Warning") {
                $allTasksTpHead = $subHead01war
            } ElseIf ($arrAllTasksTp.Status -match "Success") {
                $allTasksTpHead = $subHead01suc
            } Else {
                $allTasksTpHead = $subHead01
            }
            $bodyAllTasksTp = $allTasksTpHead + "Tape Backup Tasks" + $subHead02 + $bodyAllTasksTp
        } Else {
            $arrAllTasksTp = $taskListTp | Select-Object @{Name = "Name"; Expression = { $_.Name } },
            @{Name = "Job Name"; Expression = { $_.JobSess.Name } },
            @{Name = "Start Time"; Expression = { $_.Progress.StartTimeLocal } },
            @{Name = "Stop Time"; Expression = { If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") { "-" } Else { $_.Progress.StopTimeLocal } } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Details"; Expression = { ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } }, Status
            $bodyAllTasksTp = $arrAllTasksTp | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
            If ($arrAllTasksTp.Status -match "Failed") {
                $allTasksTpHead = $subHead01err
            } ElseIf ($arrAllTasksTp.Status -match "Warning") {
                $allTasksTpHead = $subHead01war
            } ElseIf ($arrAllTasksTp.Status -match "Success") {
                $allTasksTpHead = $subHead01suc
            } Else {
                $allTasksTpHead = $subHead01
            }
            $bodyAllTasksTp = $allTasksTpHead + "Tape Backup Tasks" + $subHead02 + $bodyAllTasksTp
        }
    }
}

# Get Pending Tape Backup Tasks
$bodyTasksPendingTp = $null
If ($showPendingTasksTp) {
    If ($pendingTasksTp.count -gt 0) {
        $bodyTasksPendingTp = $pendingTasksTp | Select-Object @{Name = "Name"; Expression = { $_.Name } },
        @{Name = "Job Name"; Expression = { $_.JobSess.Name } },
        @{Name = "Start Time"; Expression = { $_.Info.Progress.StartTimeLocal } },
        @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
        @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Progress.AvgSpeed / 1MB, 2) } },
        @{Name = "Read (GB)"; Expression = { [Math]::Round($_.Progress.ReadSize / 1GB, 2) } },
        @{Name = "Transferred (GB)"; Expression = { [Math]::Round($_.Progress.TransferedSize / 1GB, 2) } },
        Status | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
        $bodyTasksPendingTp = $subHead01 + "Pending Tape Backup Tasks" + $subHead02 + $bodyTasksPendingTp
    }
}

# Get Working Tape Backup Tasks
$bodyTasksRunningTp = $null
If ($showRunningTasksTp) {
    If ($runningTasksTp.count -gt 0) {
        $bodyTasksRunningTp = $runningTasksTp | Select-Object @{Name = "Name"; Expression = { $_.Name } },
        @{Name = "Job Name"; Expression = { $_.JobSess.Name } },
        @{Name = "Start Time"; Expression = { $_.Info.Progress.StartTimeLocal } },
        @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
        @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Progress.AvgSpeed / 1MB, 2) } },
        @{Name = "Read (GB)"; Expression = { [Math]::Round($_.Progress.ReadSize / 1GB, 2) } },
        @{Name = "Transferred (GB)"; Expression = { [Math]::Round($_.Progress.TransferedSize / 1GB, 2) } },
        Status | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
        $bodyTasksRunningTp = $subHead01 + "Working Tape Backup Tasks" + $subHead02 + $bodyTasksRunningTp
    }
}

# Get Tape Backup Tasks with Warnings or Failures
$bodyTaskWFTp = $null
If ($showTaskWFTp) {
    If ($wfTasksTp.count -gt 0) {
        If ($showDetailedTp) {
            $arrTaskWFTp = $wfTasksTp | Select-Object @{Name = "Name"; Expression = { $_.Name } },
            @{Name = "Job Name"; Expression = { $_.JobSess.Name } },
            @{Name = "Start Time"; Expression = { $_.Progress.StartTimeLocal } },
            @{Name = "Stop Time"; Expression = { $_.Progress.StopTimeLocal } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Progress.AvgSpeed / 1MB, 2) } },
            @{Name = "Total (GB)"; Expression = { [Math]::Round($_.Progress.ProcessedSize / 1GB, 2) } },
            @{Name = "Data Read (GB)"; Expression = { [Math]::Round($_.Progress.ReadSize / 1GB, 2) } },
            @{Name = "Transferred (GB)"; Expression = { [Math]::Round($_.Progress.TransferedSize / 1GB, 2) } },
            @{Name = "Details"; Expression = { ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } }, Status
            $bodyTaskWFTp = $arrTaskWFTp | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
            If ($arrTaskWFTp.Status -match "Failed") {
                $taskWFTpHead = $subHead01err
            } ElseIf ($arrTaskWFTp.Status -match "Warning") {
                $taskWFTpHead = $subHead01war
            } ElseIf ($arrTaskWFTp.Status -match "Success") {
                $taskWFTpHead = $subHead01suc
            } Else {
                $taskWFTpHead = $subHead01
            }
            $bodyTaskWFTp = $taskWFTpHead + "Tape Backup Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFTp
        } Else {
            $arrTaskWFTp = $wfTasksTp | Select-Object @{Name = "Name"; Expression = { $_.Name } },
            @{Name = "Job Name"; Expression = { $_.JobSess.Name } },
            @{Name = "Start Time"; Expression = { $_.Progress.StartTimeLocal } },
            @{Name = "Stop Time"; Expression = { $_.Progress.StopTimeLocal } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Details"; Expression = { ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } }, Status
            $bodyTaskWFTp = $arrTaskWFTp | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
            If ($arrTaskWFTp.Status -match "Failed") {
                $taskWFTpHead = $subHead01err
            } ElseIf ($arrTaskWFTp.Status -match "Warning") {
                $taskWFTpHead = $subHead01war
            } ElseIf ($arrTaskWFTp.Status -match "Success") {
                $taskWFTpHead = $subHead01suc
            } Else {
                $taskWFTpHead = $subHead01
            }
            $bodyTaskWFTp = $taskWFTpHead + "Tape Backup Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFTp
        }
    }
}

# Get Successful Tape Backup Tasks
$bodyTaskSuccTp = $null
If ($showTaskSuccessTp) {
    If ($successTasksTp.count -gt 0) {
        If ($showDetailedTp) {
            $bodyTaskSuccTp = $successTasksTp | Select-Object @{Name = "Name"; Expression = { $_.Name } },
            @{Name = "Job Name"; Expression = { $_.JobSess.Name } },
            @{Name = "Start Time"; Expression = { $_.Progress.StartTimeLocal } },
            @{Name = "Stop Time"; Expression = {
                    If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") { "-" }
                    Else { $_.Progress.StopTimeLocal }
                }
            },
            @{Name = "Duration (HH:MM:SS)"; Expression = {
                    If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") { "-" }
                    Else { Get-Duration -ts $_.Progress.Duration }
                }
            },
            @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Progress.AvgSpeed / 1MB, 2) } },
            @{Name = "Total (GB)"; Expression = { [Math]::Round($_.Progress.ProcessedSize / 1GB, 2) } },
            @{Name = "Data Read (GB)"; Expression = { [Math]::Round($_.Progress.ReadSize / 1GB, 2) } },
            @{Name = "Transferred (GB)"; Expression = { [Math]::Round($_.Progress.TransferedSize / 1GB, 2) } },
            Status | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
            $bodyTaskSuccTp = $subHead01suc + "Successful Tape Backup Tasks" + $subHead02 + $bodyTaskSuccTp
        } Else {
            $bodyTaskSuccTp = $successTasksTp | Select-Object @{Name = "Name"; Expression = { $_.Name } },
            @{Name = "Job Name"; Expression = { $_.JobSess.Name } },
            @{Name = "Start Time"; Expression = { $_.Progress.StartTimeLocal } },
            @{Name = "Stop Time"; Expression = {
                    If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") { "-" }
                    Else { $_.Progress.StopTimeLocal }
                }
            },
            @{Name = "Duration (HH:MM:SS)"; Expression = {
                    If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") { "-" }
                    Else { Get-Duration -ts $_.Progress.Duration }
                }
            },
            Status | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
            $bodyTaskSuccTp = $subHead01suc + "Successful Tape Backup Tasks" + $subHead02 + $bodyTaskSuccTp
        }
    }
}

# Get all Tapes
$bodyTapes = $null
If ($showTapes) {
    $expTapes = @($mediaTapes)
    if ($expTapes.Count -gt 0) {
        $expTapes = $expTapes | Select-Object Name, Barcode,
        @{Name = "Media Pool"; Expression = {
                $poolId = $_.MediaPoolId
                ($mediaPools | Where-Object { $_.Id -eq $poolId }).Name
            }
        },
        @{Name = "Media Set"; Expression = { $_.MediaSet } }, @{Name = "Sequence #"; Expression = { $_.SequenceNumber } },
        @{Name = "Location"; Expression = {
                switch ($_.Location) {
                    "None" { "Offline" }
                    "Slot" {
                        $lId = $_.LibraryId
                        $lName = $($mediaLibs | Where-Object { $_.Id -eq $lId }).Name
                        [int]$slot = $_.SlotAddress + 1
                        "{0} : {1} {2}" -f $lName, $_, $slot
                    }
                    "Drive" {
                        $lId = $_.LibraryId
                        $dId = $_.DriveId
                        $lName = $($mediaLibs | Where-Object { $_.Id -eq $lId }).Name
                        $dName = $($mediaDrives | Where-Object { $_.Id -eq $dId }).Name
                        [int]$dNum = $_.Location.DriveAddress + 1
                        "{0} : {1} {2} (Drive ID: {3})" -f $lName, $_, $dNum, $dName
                    }
                    "Vault" {
                        $vId = $_.VaultId
                        $vName = $($mediaVaults | Where-Object { $_.Id -eq $vId }).Name
                        "{0}: {1}" -f $_, $vName
                    }
                    default { "Lost in Space" }
                }
            }
        },
        @{Name = "Capacity (GB)"; Expression = { [Math]::Round([Decimal]$_.Capacity / 1GB, 2) } },
        @{Name = "Free (GB)"; Expression = { [Math]::Round([Decimal]$_.Free / 1GB, 2) } },
        @{Name = "Last Write"; Expression = { $_.LastWriteTime } },
        @{Name = "Expiration Date"; Expression = {
                If ($(Get-Date $_.ExpirationDate) -lt $(Get-Date)) {
                    "Expired"
                } Else {
                    $_.ExpirationDate
                }
            }
        } | Sort-Object Name | ConvertTo-HTML -Fragment
        $bodyTapes = $subHead01 + "All Tapes" + $subHead02 + $expTapes
    }
}

# Get all Tapes in each Custom Media Pool
$bodyTpPool = $null
If ($showTpMp) {
    ForEach ($mp in ($mediaPools | Where-Object { $_.Type -eq "Custom" } | Sort-Object Name)) {
        $expTapes = @($mediaTapes | Where-Object { ($_.MediaPoolId -eq $mp.Id) })
        if ($expTapes.Count -gt 0) {
            $expTapes = $expTapes | Select-Object Name, Barcode,
            @{Name = "Media Set"; Expression = { $_.MediaSet } }, @{Name = "Sequence #"; Expression = { $_.SequenceNumber } },
            @{Name = "Location"; Expression = {
                    switch ($_.Location) {
                        "None" { "Offline" }
                        "Slot" {
                            $lId = $_.LibraryId
                            $lName = $($mediaLibs | Where-Object { $_.Id -eq $lId }).Name
                            [int]$slot = $_.SlotAddress + 1
                            "{0} : {1} {2}" -f $lName, $_, $slot
                        }
                        "Drive" {
                            $lId = $_.LibraryId
                            $dId = $_.DriveId
                            $lName = $($mediaLibs | Where-Object { $_.Id -eq $lId }).Name
                            $dName = $($mediaDrives | Where-Object { $_.Id -eq $dId }).Name
                            [int]$dNum = $_.Location.DriveAddress + 1
                            "{0} : {1} {2} (Drive ID: {3})" -f $lName, $_, $dNum, $dName
                        }
                        "Vault" {
                            $vId = $_.VaultId
                            $vName = $($mediaVaults | Where-Object { $_.Id -eq $vId }).Name
                            "{0}: {1}" -f $_, $vName
                        }
                        default { "Lost in Space" }
                    }
                }
            },
            @{Name = "Capacity (GB)"; Expression = { [Math]::Round([Decimal]$_.Capacity / 1GB, 2) } },
            @{Name = "Free (GB)"; Expression = { [Math]::Round([Decimal]$_.Free / 1GB, 2) } },
            @{Name = "Last Write"; Expression = { $_.LastWriteTime } },
            @{Name = "Expiration Date"; Expression = {
                    If ($(Get-Date $_.ExpirationDate) -lt $(Get-Date)) {
                        "Expired"
                    } Else {
                        $_.ExpirationDate
                    }
                }
            } | Sort-Object "Last Write" | ConvertTo-HTML -Fragment
            $bodyTpPool += $subHead01 + "All Tapes in Media Pool: " + $mp.Name + $subHead02 + $expTapes
        }
    }
}

# Get all Tapes in each Vault
$bodyTpVlt = $null
If ($showTpVlt) {
    ForEach ($vlt in ($mediaVaults | Sort-Object Name)) {
        $expTapes = @($mediaTapes | Where-Object { ($_.Location.VaultId -eq $vlt.Id) })
        if ($expTapes.Count -gt 0) {
            $expTapes = $expTapes | Select-Object Name, Barcode,
            @{Name = "Media Pool"; Expression = {
                    $poolId = $_.MediaPoolId
                    ($mediaPools | Where-Object { $_.Id -eq $poolId }).Name
                }
            },
            @{Name = "Media Set"; Expression = { $_.MediaSet } }, @{Name = "Sequence #"; Expression = { $_.SequenceNumber } },
            @{Name = "Capacity (GB)"; Expression = { [Math]::Round([Decimal]$_.Capacity / 1GB, 2) } },
            @{Name = "Free (GB)"; Expression = { [Math]::Round([Decimal]$_.Free / 1GB, 2) } },
            @{Name = "Last Write"; Expression = { $_.LastWriteTime } },
            @{Name = "Expiration Date"; Expression = {
                    If ($(Get-Date $_.ExpirationDate) -lt $(Get-Date)) {
                        "Expired"
                    } Else {
                        $_.ExpirationDate
                    }
                }
            } | Sort-Object Name | ConvertTo-HTML -Fragment
            $bodyTpVlt += $subHead01 + "All Tapes in Vault: " + $vlt.Name + $subHead02 + $expTapes
        }
    }
}

# Get all Expired Tapes
$bodyExpTp = $null
If ($showExpTp) {
    $expTapes = @($mediaTapes | Where-Object { ($_.IsExpired -eq $True) })
    if ($expTapes.Count -gt 0) {
        $expTapes = $expTapes | Select-Object Name, Barcode,
        @{Name = "Media Pool"; Expression = {
                $poolId = $_.MediaPoolId
                ($mediaPools | Where-Object { $_.Id -eq $poolId }).Name
            }
        },
        @{Name = "Media Set"; Expression = { $_.MediaSet } }, @{Name = "Sequence #"; Expression = { $_.SequenceNumber } },
        @{Name = "Location"; Expression = {
                switch ($_.Location) {
                    "None" { "Offline" }
                    "Slot" {
                        $lId = $_.LibraryId
                        $lName = $($mediaLibs | Where-Object { $_.Id -eq $lId }).Name
                        [int]$slot = $_.SlotAddress + 1
                        "{0} : {1} {2}" -f $lName, $_, $slot
                    }
                    "Drive" {
                        $lId = $_.LibraryId
                        $dId = $_.DriveId
                        $lName = $($mediaLibs | Where-Object { $_.Id -eq $lId }).Name
                        $dName = $($mediaDrives | Where-Object { $_.Id -eq $dId }).Name
                        [int]$dNum = $_.Location.DriveAddress + 1
                        "{0} : {1} {2} (Drive ID: {3})" -f $lName, $_, $dNum, $dName
                    }
                    "Vault" {
                        $vId = $_.VaultId
                        $vName = $($mediaVaults | Where-Object { $_.Id -eq $vId }).Name
                        "{0}: {1}" -f $_, $vName
                    }
                    default { "Lost in Space" }
                }
            }
        },
        @{Name = "Capacity (GB)"; Expression = { [Math]::Round([Decimal]$_.Capacity / 1GB, 2) } },
        @{Name = "Free (GB)"; Expression = { [Math]::Round([Decimal]$_.Free / 1GB, 2) } },
        @{Name = "Last Write"; Expression = { $_.LastWriteTime } } | Sort-Object Name | ConvertTo-HTML -Fragment
        $bodyExpTp = $subHead01 + "All Expired Tapes" + $subHead02 + $expTapes
    }
}

# Get Expired Tapes in each Custom Media Pool
$bodyTpExpPool = $null
If ($showExpTpMp) {
    ForEach ($mp in ($mediaPools | Where-Object { $_.Type -eq "Custom" } | Sort-Object Name)) {
        $expTapes = @($mediaTapes | Where-Object { ($_.MediaPoolId -eq $mp.Id -and $_.IsExpired -eq $True) })
        if ($expTapes.Count -gt 0) {
            $expTapes = $expTapes | Select-Object Name, Barcode,
            @{Name = "Media Set"; Expression = { $_.MediaSet } }, @{Name = "Sequence #"; Expression = { $_.SequenceNumber } },
            @{Name = "Location"; Expression = {
                    switch ($_.Location) {
                        "None" { "Offline" }
                        "Slot" {
                            $lId = $_.LibraryId
                            $lName = $($mediaLibs | Where-Object { $_.Id -eq $lId }).Name
                            [int]$slot = $_.SlotAddress + 1
                            "{0} : {1} {2}" -f $lName, $_, $slot
                        }
                        "Drive" {
                            $lId = $_.LibraryId
                            $dId = $_.DriveId
                            $lName = $($mediaLibs | Where-Object { $_.Id -eq $lId }).Name
                            $dName = $($mediaDrives | Where-Object { $_.Id -eq $dId }).Name
                            [int]$dNum = $_.Location.DriveAddress + 1
                            "{0} : {1} {2} (Drive ID: {3})" -f $lName, $_, $dNum, $dName
                        }
                        "Vault" {
                            $vId = $_.VaultId
                            $vName = $($mediaVaults | Where-Object { $_.Id -eq $vId }).Name
                            "{0}: {1}" -f $_, $vName
                        }
                        default { "Lost in Space" }
                    }
                }
            },
            @{Name = "Capacity (GB)"; Expression = { [Math]::Round([Decimal]$_.Capacity / 1GB, 2) } },
            @{Name = "Free (GB)"; Expression = { [Math]::Round([Decimal]$_.Free / 1GB, 2) } },
            @{Name = "Last Write"; Expression = { $_.LastWriteTime } } | Sort-Object "Last Write" | ConvertTo-HTML -Fragment
            $bodyTpExpPool += $subHead01 + "Expired Tapes in Media Pool: " + $mp.Name + $subHead02 + $expTapes
        }
    }
}

# Get Expired Tapes in each Vault
$bodyTpExpVlt = $null
If ($showExpTpVlt) {
    ForEach ($vlt in ($mediaVaults | Sort-Object Name)) {
        $expTapes = @($mediaTapes | Where-Object { ($_.Location.VaultId -eq $vlt.Id -and $_.IsExpired -eq $True) })
        if ($expTapes.Count -gt 0) {
            $expTapes = $expTapes | Select-Object Name, Barcode,
            @{Name = "Media Pool"; Expression = {
                    $poolId = $_.MediaPoolId
                    ($mediaPools | Where-Object { $_.Id -eq $poolId }).Name
                }
            },
            @{Name = "Media Set"; Expression = { $_.MediaSet } }, @{Name = "Sequence #"; Expression = { $_.SequenceNumber } },
            @{Name = "Capacity (GB)"; Expression = { [Math]::Round([Decimal]$_.Capacity / 1GB, 2) } },
            @{Name = "Free (GB)"; Expression = { [Math]::Round([Decimal]$_.Free / 1GB, 2) } },
            @{Name = "Last Write"; Expression = { $_.LastWriteTime } } | Sort-Object "Last Write" | ConvertTo-HTML -Fragment
            $bodyTpExpVlt += $subHead01 + "Expired Tapes in Vault: " + $vlt.Name + $subHead02 + $expTapes
        }
    }
}

# Get all Tapes written to within time frame
$bodyTpWrt = $null
If ($showTpWrt) {
    $expTapes = @($mediaTapes | Where-Object { $_.LastWriteTime -ge (Get-Date).AddHours(-$HourstoCheck) })
    if ($expTapes.Count -gt 0) {
        $expTapes = $expTapes | Select-Object Name, Barcode,
        @{Name = "Media Pool"; Expression = {
                $poolId = $_.MediaPoolId
                ($mediaPools | Where-Object { $_.Id -eq $poolId }).Name
            }
        },
        @{Name = "Media Set"; Expression = { $_.MediaSet } }, @{Name = "Sequence #"; Expression = { $_.SequenceNumber } },
        @{Name = "Location"; Expression = {
                switch ($_.Location) {
                    "None" { "Offline" }
                    "Slot" {
                        $lId = $_.LibraryId
                        $lName = $($mediaLibs | Where-Object { $_.Id -eq $lId }).Name
                        [int]$slot = $_.SlotAddress + 1
                        "{0} : {1} {2}" -f $lName, $_, $slot
                    }
                    "Drive" {
                        $lId = $_.LibraryId
                        $dId = $_.DriveId
                        $lName = $($mediaLibs | Where-Object { $_.Id -eq $lId }).Name
                        $dName = $($mediaDrives | Where-Object { $_.Id -eq $dId }).Name
                        [int]$dNum = $_.Location.DriveAddress + 1
                        "{0} : {1} {2} (Drive ID: {3})" -f $lName, $_, $dNum, $dName
                    }
                    "Vault" {
                        $vId = $_.VaultId
                        $vName = $($mediaVaults | Where-Object { $_.Id -eq $vId }).Name
                        "{0}: {1}" -f $_, $vName
                    }
                    default { "Lost in Space" }
                }
            }
        },
        @{Name = "Capacity (GB)"; Expression = { [Math]::Round([Decimal]$_.Capacity / 1GB, 2) } },
        @{Name = "Free (GB)"; Expression = { [Math]::Round([Decimal]$_.Free / 1GB, 2) } },
        @{Name = "Last Write"; Expression = { $_.LastWriteTime } },
        @{Name = "Expiration Date"; Expression = {
                If ($(Get-Date $_.ExpirationDate) -lt $(Get-Date)) {
                    "Expired"
                } Else {
                    $_.ExpirationDate
                }
            }
        } | Sort-Object "Last Write" | ConvertTo-HTML -Fragment
        $bodyTpWrt = $subHead01 + "All Tapes Written" + $subHead02 + $expTapes
    }
}

# Get Agent Backup Summary Info
$bodySummaryEp = $null
If ($showSummaryEp) {
    $vbrEpHash = @{
        "Sessions"   = If ($sessListEp) { @($sessListEp).Count } Else { 0 }
        "Successful" = @($successSessionsEp).Count
        "Warning"    = @($warningSessionsEp).Count
        "Fails"      = @($failsSessionsEp).Count
        "Running"    = @($runningSessionsEp).Count
    }
    $vbrEPObj = New-Object -TypeName PSObject -Property $vbrEpHash
    If ($onlyLastEp) {
        $total = "Jobs Run"
    } Else {
        $total = "Total Sessions"
    }
    $arrSummaryEp = $vbrEPObj | Select-Object @{Name = $total; Expression = { $_.Sessions } },
    @{Name = "Running"; Expression = { $_.Running } }, @{Name = "Successful"; Expression = { $_.Successful } },
    @{Name = "Warnings"; Expression = { $_.Warning } }, @{Name = "Failures"; Expression = { $_.Fails } }
    $bodySummaryEp = $arrSummaryEp | ConvertTo-HTML -Fragment
    If ($arrSummaryEp.Failures -gt 0) {
        $summaryEpHead = $subHead01err
    } ElseIf ($arrSummaryEp.Warnings -gt 0) {
        $summaryEpHead = $subHead01war
    } ElseIf ($arrSummaryEp.Successful -gt 0) {
        $summaryEpHead = $subHead01suc
    } Else {
        $summaryEpHead = $subHead01
    }
    $bodySummaryEp = $summaryEpHead + "Agent Backup Results Summary" + $subHead02 + $bodySummaryEp
}

# Get Agent Backup Job Status
$bodyJobsEp = $null
If ($showJobsEp) {
    If ($allJobsEp.count -gt 0) {
        $bodyJobsEp = $allJobsEp | Sort-Object Name | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
        @{Name = "Description"; Expression = { $_.Description } },
        @{Name = "Enabled"; Expression = { $_.IsEnabled } }, @{Name = "Status"; Expression = { $_.LastState } },
        @{Name = "Target Repo"; Expression = { $_.Target } }, @{Name = "Next Run"; Expression = { $_.NextRun } },
        @{Name = "Last Result"; Expression = { If ($_.LastResult -eq "None") { "" }Else { $_.LastResult } } } | ConvertTo-HTML -Fragment
        $bodyJobsEp = $subHead01 + "Agent Backup Job Status" + $subHead02 + $bodyJobsEp
    }
}

# Get Agent Backup Job Size
$bodyJobSizeEp = $null
If ($showBackupSizeEp) {
    If ($backupsEp.count -gt 0) {
        $bodyJobSizeEp = Get-BackupSize -backups $backupsEp | Sort-Object JobName | Select-Object @{Name = "Job Name"; Expression = { $_.JobName } },
        @{Name = "VM Count"; Expression = { $_.VMCount } },
        @{Name = "Repository"; Expression = { $_.Repo } },
        @{Name = "Data Size (GB)"; Expression = { $_.DataSize } },
        @{Name = "Backup Size (GB)"; Expression = { $_.BackupSize } } | ConvertTo-HTML -Fragment
        $bodyJobSizeEp = $subHead01 + "Agent Backup Job Size" + $subHead02 + $bodyJobSizeEp
    }
}

# Get Agent Backup Sessions
$bodyAllSessEp = @()
$arrAllSessEp = @()
If ($showAllSessEp) {
    If ($sessListEp.count -gt 0) {
        Foreach ($job in $allJobsEp) {
            $arrAllSessEp += $sessListEp | Where-Object { $_.JobId -eq $job.Id } | Select-Object @{Name = "Job Name"; Expression = { $job.Name } },
            @{Name = "State"; Expression = { $_.State } }, @{Name = "Start Time"; Expression = { $_.CreationTime } },
            @{Name = "Stop Time"; Expression = { If ($_.EndTime -eq "1/1/1900 12:00:00 AM") { "-" } Else { $_.EndTime } } },
            @{Name = "Duration (HH:MM:SS)"; Expression = {
                    If ($_.EndTime -eq "1/1/1900 12:00:00 AM") {
                        Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))
                    } Else {
                        Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime)
                    }
                }
            }, Result
        }
        $bodyAllSessEp = $arrAllSessEp | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
        If ($arrAllSessEp.Result -match "Failed") {
            $allSessEpHead = $subHead01err
        } ElseIf ($arrAllSessEp.Result -match "Warning") {
            $allSessEpHead = $subHead01war
        } ElseIf ($arrAllSessEp.Result -match "Success") {
            $allSessEpHead = $subHead01suc
        } Else {
            $allSessEpHead = $subHead01
        }
        $bodyAllSessEp = $allSessEpHead + "Agent Backup Sessions" + $subHead02 + $bodyAllSessEp
    }
}

# Get Running Agent Backup Jobs
$bodyRunningEp = @()
If ($showRunningEp) {
    If ($runningSessionsEp.count -gt 0) {
        Foreach ($job in $allJobsEp) {
            $bodyRunningEp += $runningSessionsEp | Where-Object { $_.JobId -eq $job.Id } | Select-Object @{Name = "Job Name"; Expression = { $job.Name } },
            @{Name = "Start Time"; Expression = { $_.CreationTime } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date)) } }
        }
        $bodyRunningEp = $bodyRunningEp | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
        $bodyRunningEp = $subHead01 + "Running Agent Backup Jobs" + $subHead02 + $bodyRunningEp
    }
}

# Get Agent Backup Sessions with Warnings or Failures
$bodySessWFEp = @()
$arrSessWFEp = @()
If ($showWarnFailEp) {
    $sessWFEp = @($warningSessionsEp + $failsSessionsEp)
    If ($sessWFEp.count -gt 0) {
        If ($onlyLastEp) {
            $headerWFEp = "Agent Backup Jobs with Warnings or Failures"
        } Else {
            $headerWFEp = "Agent Backup Sessions with Warnings or Failures"
        }
        Foreach ($job in $allJobsEp) {
            $arrSessWFEp += $sessWFEp | Where-Object { $_.JobId -eq $job.Id } | Select-Object @{Name = "Job Name"; Expression = { $job.Name } },
            @{Name = "Start Time"; Expression = { $_.CreationTime } }, @{Name = "Stop Time"; Expression = { $_.EndTime } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime) } },
            Result
        }
        $bodySessWFEp = $arrSessWFEp | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
        If ($arrSessWFEp.Result -match "Failed") {
            $sessWFEpHead = $subHead01err
        } ElseIf ($arrSessWFEp.Result -match "Warning") {
            $sessWFEpHead = $subHead01war
        } ElseIf ($arrSessWFEp.Result -match "Success") {
            $sessWFEpHead = $subHead01suc
        } Else {
            $sessWFEpHead = $subHead01
        }
        $bodySessWFEp = $sessWFEpHead + $headerWFEp + $subHead02 + $bodySessWFEp
    }
}

# Get Successful Agent Backup Sessions
$bodySessSuccEp = @()
If ($showSuccessEp) {
    If ($successSessionsEp.count -gt 0) {
        If ($onlyLastEp) {
            $headerSuccEp = "Successful Agent Backup Jobs"
        } Else {
            $headerSuccEp = "Successful Agent Backup Sessions"
        }
        Foreach ($job in $allJobsEp) {
            $bodySessSuccEp += $successSessionsEp | Where-Object { $_.JobId -eq $job.Id } | Select-Object @{Name = "Job Name"; Expression = { $job.Name } },
            @{Name = "Start Time"; Expression = { $_.CreationTime } }, @{Name = "Stop Time"; Expression = { $_.EndTime } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime) } },
            Result
        }
        $bodySessSuccEp = $bodySessSuccEp | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
        $bodySessSuccEp = $subHead01suc + $headerSuccEp + $subHead02 + $bodySessSuccEp
    }
}

# Get SureBackup Summary Info
$bodySummarySb = $null
If ($showSummarySb) {
    $vbrMasterHash = @{
        "Sessions"   = If ($sessListSb) { @($sessListSb).Count } Else { 0 }
        "Successful" = @($successSessionsSb).Count
        "Warning"    = @($warningSessionsSb).Count
        "Fails"      = @($failsSessionsSb).Count
        "Running"    = @($runningSessionsSb).Count
    }
    $vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
    If ($onlyLastSb) {
        $total = "Jobs Run"
    } Else {
        $total = "Total Sessions"
    }
    $arrSummarySb = $vbrMasterObj | Select-Object @{Name = $total; Expression = { $_.Sessions } },
    @{Name = "Running"; Expression = { $_.Running } }, @{Name = "Successful"; Expression = { $_.Successful } },
    @{Name = "Warnings"; Expression = { $_.Warning } }, @{Name = "Failures"; Expression = { $_.Fails } }
    $bodySummarySb = $arrSummarySb | ConvertTo-HTML -Fragment
    If ($arrSummarySb.Failures -gt 0) {
        $summarySbHead = $subHead01err
    } ElseIf ($arrSummarySb.Warnings -gt 0) {
        $summarySbHead = $subHead01war
    } ElseIf ($arrSummarySb.Successful -gt 0) {
        $summarySbHead = $subHead01suc
    } Else {
        $summarySbHead = $subHead01
    }
    $bodySummarySb = $summarySbHead + "SureBackup Results Summary" + $subHead02 + $bodySummarySb
}

# Get SureBackup Job Status
$bodyJobsSb = $null
If ($showJobsSb) {
    If ($allJobsSb.count -gt 0) {
        $bodyJobsSb = @()
        Foreach ($SbJob in $allJobsSb) {
            $bodyJobsSb += $SbJob | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
            @{Name = "Enabled"; Expression = { $_.IsScheduleEnabled } },
            @{Name = "Status"; Expression = {
                    If ($_.GetLastState() -eq "Working") {
                        $currentSess = $_.FindLastSession()
                        $csessPercent = $currentSess.CompletionPercentage
                        $cStatus = "$($csessPercent)% completed"
                        $cStatus
                    } Else {
                        $_.GetLastState()
                    }
                }
            },
            @{Name = "Virtual Lab"; Expression = { $(Get-VSBVirtualLab | Where-Object { $_.Id -eq $SbJob.VirtualLabId }).Name } },
            @{Name = "Linked Jobs"; Expression = { $($_.GetLinkedJobs()).Name -join "," } },
            @{Name = "Next Run"; Expression = {
                    If ($_.IsScheduleEnabled -eq $false) { "<Disabled>" }
                    ElseIf ($_.JobOptions.RunManually) { "<not scheduled>" }
                    ElseIf ($_.ScheduleOptions.IsContinious) { "<Continious>" }
                    ElseIf ($_.ScheduleOptions.OptionsScheduleAfterJob.IsEnabled) { "After [" + $(($allJobs + $allJobsTp) | Where-Object { $_.Id -eq $SbJob.Info.ParentScheduleId }).Name + "]" }
                    Else { $_.ScheduleOptions.NextRun } }
            },
            @{Name = "Last Result"; Expression = { If ($_.GetLastResult() -eq "None") { "" }Else { $_.GetLastResult() } } }
        }
        $bodyJobsSb = $bodyJobsSb | Sort-Object "Next Run" | ConvertTo-HTML -Fragment
        $bodyJobsSb = $subHead01 + "SureBackup Job Status" + $subHead02 + $bodyJobsSb
    }
}

# Get SureBackup Sessions
$bodyAllSessSb = $null
If ($showAllSessSb) {
    If ($sessListSb.count -gt 0) {
        $arrAllSessSb = $sessListSb | Sort-Object Creationtime | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
        @{Name = "State"; Expression = { $_.State } },
        @{Name = "Start Time"; Expression = { $_.CreationTime } },
        @{Name = "Stop Time"; Expression = { If ($_.EndTime -eq "1/1/1900 12:00:00 AM") { "-" } Else { $_.EndTime } } },

        @{Name = "Duration (HH:MM:SS)"; Expression = {
                If ($_.EndTime -eq "1/1/1900 12:00:00 AM") {
                    Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))
                } Else {
                    Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime)
                }
            }
        }, Result
        $bodyAllSessSb = $arrAllSessSb | ConvertTo-HTML -Fragment
        If ($arrAllSessSb.Result -match "Failed") {
            $allSessSbHead = $subHead01err
        } ElseIf ($arrAllSessSb.Result -match "Warning") {
            $allSessSbHead = $subHead01war
        } ElseIf ($arrAllSessSb.Result -match "Success") {
            $allSessSbHead = $subHead01suc
        } Else {
            $allSessSbHead = $subHead01
        }
        $bodyAllSessSb = $allSessSbHead + "SureBackup Sessions" + $subHead02 + $bodyAllSessSb
    }
}

# Get Running SureBackup Jobs
$bodyRunningSb = $null
If ($showRunningSb) {
    If ($runningSessionsSb.count -gt 0) {
        $bodyRunningSb = $runningSessionsSb | Sort-Object Creationtime | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
        @{Name = "Start Time"; Expression = { $_.CreationTime } },
        @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date)) } },
        @{Name = "% Complete"; Expression = { $_.Progress } } | ConvertTo-HTML -Fragment
        $bodyRunningSb = $subHead01 + "Running SureBackup Jobs" + $subHead02 + $bodyRunningSb
    }
}

# Get SureBackup Sessions with Warnings or Failures
$bodySessWFSb = $null
If ($showWarnFailSb) {
    $sessWF = @($warningSessionsSb + $failsSessionsSb)
    If ($sessWF.count -gt 0) {
        If ($onlyLastSb) {
            $headerWF = "SureBackup Jobs with Warnings or Failures"
        } Else {
            $headerWF = "SureBackup Sessions with Warnings or Failures"
        }
        $arrSessWFSb = $sessWF | Sort-Object Creationtime | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
        @{Name = "Start Time"; Expression = { $_.CreationTime } },
        @{Name = "Stop Time"; Expression = { $_.EndTime } },
        @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime) } }, Result
        $bodySessWFSb = $arrSessWFSb | ConvertTo-HTML -Fragment
        If ($arrSessWFSb.Result -match "Failed") {
            $sessWFSbHead = $subHead01err
        } ElseIf ($arrSessWFSb.Result -match "Warning") {
            $sessWFSbHead = $subHead01war
        } ElseIf ($arrSessWFSb.Result -match "Success") {
            $sessWFSbHead = $subHead01suc
        } Else {
            $sessWFSbHead = $subHead01
        }
        $bodySessWFSb = $sessWFSbHead + $headerWF + $subHead02 + $bodySessWFSb
    }
}

# Get Successful SureBackup Sessions
$bodySessSuccSb = $null
If ($showSuccessSb) {
    If ($successSessionsSb.count -gt 0) {
        If ($onlyLastSb) {
            $headerSucc = "Successful SureBackup Jobs"
        } Else {
            $headerSucc = "Successful SureBackup Sessions"
        }
        $bodySessSuccSb = $successSessionsSb | Sort-Object Creationtime | Select-Object @{Name = "Job Name"; Expression = { $_.Name } },
        @{Name = "Start Time"; Expression = { $_.CreationTime } },
        @{Name = "Stop Time"; Expression = { $_.EndTime } },
        @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime) } },
        Result | ConvertTo-HTML -Fragment
        $bodySessSuccSb = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccSb
    }
}

## Gathering tasks after session info has been recorded due to Veeam issue
# Gather all SureBackup Tasks from Sessions within time frame
$taskListSb = @()
$taskListSb += $sessListSb | Get-VSBTaskSession
$successTasksSb = @($taskListSb | Where-Object { $_.Info.Result -eq "Success" })
$wfTasksSb = @($taskListSb | Where-Object { $_.Info.Result -match "Warning|Failed" })
$runningTasksSb = @()
$runningTasksSb += $runningSessionsSb | Get-VSBTaskSession | Where-Object { $_.Status -ne "Stopped" }

# Get SureBackup Tasks
$bodyAllTasksSb = $null
If ($showAllTasksSb) {
    If ($taskListSb.count -gt 0) {
        $arrAllTasksSb = $taskListSb | Select-Object @{Name = "VM Name"; Expression = { $_.Name } },
        @{Name = "Job Name"; Expression = { $_.JobSession.JobName } },
        @{Name = "Status"; Expression = { $_.Status } },
        @{Name = "Start Time"; Expression = { $_.Info.StartTime } },
        @{Name = "Stop Time"; Expression = { If ($_.Info.FinishTime -eq "1/1/1900 12:00:00 AM") { "-" } Else { $_.Info.FinishTime } } },
        @{Name = "Duration (HH:MM:SS)"; Expression = {
                If ($_.Info.FinishTime -eq "1/1/1900 12:00:00 AM") {
                    Get-Duration -ts $(New-TimeSpan $_.Info.StartTime $(Get-Date))
                } Else {
                    Get-Duration -ts $(New-TimeSpan $_.Info.StartTime $_.Info.FinishTime)
                }
            }
        },
        @{Name = "Heartbeat Test"; Expression = { $_.HeartbeatStatus } },
        @{Name = "Ping Test"; Expression = { $_.PingStatus } },
        @{Name = "Script Test"; Expression = { $_.TestScriptStatus } },
        @{Name = "Validation Test"; Expression = { $_.VadiationTestStatus } },
        @{Name = "Result"; Expression = {
                If ($_.Info.Result -eq "notrunning") {
                    "None"
                } Else {
                    $_.Info.Result
                }
            }
        }
        $bodyAllTasksSb = $arrAllTasksSb | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
        If ($arrAllTasksSb.Result -match "Failed") {
            $allTasksSbHead = $subHead01err
        } ElseIf ($arrAllTasksSb.Result -match "Warning") {
            $allTasksSbHead = $subHead01war
        } ElseIf ($arrAllTasksSb.Result -match "Success") {
            $allTasksSbHead = $subHead01suc
        } Else {
            $allTasksSbHead = $subHead01
        }
        $bodyAllTasksSb = $allTasksSbHead + "SureBackup Tasks" + $subHead02 + $bodyAllTasksSb
    }
}

# Get Running SureBackup Tasks
$bodyTasksRunningSb = $null
If ($showRunningTasksSb) {
    If ($runningTasksSb.count -gt 0) {
        $bodyTasksRunningSb = $runningTasksSb | Select-Object @{Name = "VM Name"; Expression = { $_.Name } },
        @{Name = "Job Name"; Expression = { $_.JobSession.JobName } },
        @{Name = "Start Time"; Expression = { $_.Info.StartTime } },
        @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $(New-TimeSpan $_.Info.StartTime $(Get-Date)) } },
        @{Name = "Heartbeat Test"; Expression = { $_.HeartbeatStatus } },
        @{Name = "Ping Test"; Expression = { $_.PingStatus } },
        @{Name = "Script Test"; Expression = { $_.TestScriptStatus } },
        @{Name = "Validation Test"; Expression = { $_.VadiationTestStatus } },
        Status | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
        $bodyTasksRunningSb = $subHead01 + "Running SureBackup Tasks" + $subHead02 + $bodyTasksRunningSb
    }
}

# Get SureBackup Tasks with Warnings or Failures
$bodyTaskWFSb = $null
If ($showTaskWFSb) {
    If ($wfTasksSb.count -gt 0) {
        $arrTaskWFSb = $wfTasksSb | Select-Object @{Name = "VM Name"; Expression = { $_.Name } },
        @{Name = "Job Name"; Expression = { $_.JobSession.JobName } },
        @{Name = "Start Time"; Expression = { $_.Info.StartTime } },
        @{Name = "Stop Time"; Expression = { $_.Info.FinishTime } },
        @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $(New-TimeSpan $_.Info.StartTime $_.Info.FinishTime) } },
        @{Name = "Heartbeat Test"; Expression = { $_.HeartbeatStatus } },
        @{Name = "Ping Test"; Expression = { $_.PingStatus } },
        @{Name = "Script Test"; Expression = { $_.TestScriptStatus } },
        @{Name = "Validation Test"; Expression = { $_.VadiationTestStatus } },
        @{Name = "Result"; Expression = { $_.Info.Result } }
        $bodyTaskWFSb = $arrTaskWFSb | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
        If ($arrTaskWFSb.Result -match "Failed") {
            $taskWFSbHead = $subHead01err
        } ElseIf ($arrTaskWFSb.Result -match "Warning") {
            $taskWFSbHead = $subHead01war
        } ElseIf ($arrTaskWFSb.Result -match "Success") {
            $taskWFSbHead = $subHead01suc
        } Else {
            $taskWFSbHead = $subHead01
        }
        $bodyTaskWFSb = $taskWFSbHead + "SureBackup Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFSb
    }
}

# Get Successful SureBackup Tasks
$bodyTaskSuccSb = $null
If ($showTaskSuccessSb) {
    If ($successTasksSb.count -gt 0) {
        $bodyTaskSuccSb = $successTasksSb | Select-Object @{Name = "VM Name"; Expression = { $_.Name } },
        @{Name = "Job Name"; Expression = { $_.JobSession.JobName } },
        @{Name = "Start Time"; Expression = { $_.Info.StartTime } },
        @{Name = "Stop Time"; Expression = { $_.Info.FinishTime } },
        @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $(New-TimeSpan $_.Info.StartTime $_.Info.FinishTime) } },
        @{Name = "Heartbeat Test"; Expression = { $_.HeartbeatStatus } },
        @{Name = "Ping Test"; Expression = { $_.PingStatus } },
        @{Name = "Script Test"; Expression = { $_.TestScriptStatus } },
        @{Name = "Validation Test"; Expression = { $_.VadiationTestStatus } },
        @{Name = "Result"; Expression = { $_.Info.Result } } | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
        $bodyTaskSuccSb = $subHead01suc + "Successful SureBackup Tasks" + $subHead02 + $bodyTaskSuccSb
    }
}

# Get Configuration Backup Summary Info
$bodySummaryConfig = $null
If ($showSummaryConfig) {
    $vbrConfigHash = @{
        "Enabled"        = $configBackup.Enabled
        "Status"         = $configBackup.LastState
        "Target"         = $configBackup.Target
        "Schedule"       = $configBackup.ScheduleOptions
        "Restore Points" = $configBackup.RestorePointsToKeep
        "Encrypted"      = $configBackup.EncryptionOptions.Enabled
        "Last Result"    = $configBackup.LastResult
        "Next Run"       = $configBackup.NextRun
    }
    $vbrConfigObj = New-Object -TypeName PSObject -Property $vbrConfigHash
    $bodySummaryConfig = $vbrConfigObj | Select-Object Enabled, Status, Target, Schedule, "Restore Points", "Next Run", Encrypted, "Last Result" | ConvertTo-HTML -Fragment
    If ($configBackup.LastResult -eq "Warning" -or !$configBackup.Enabled) {
        $configHead = $subHead01war
    } ElseIf ($configBackup.LastResult -eq "Success") {
        $configHead = $subHead01suc
    } ElseIf ($configBackup.LastResult -eq "Failed") {
        $configHead = $subHead01err
    } Else {
        $configHead = $subHead01
    }
    $bodySummaryConfig = $configHead + "Configuration Backup Status" + $subHead02 + $bodySummaryConfig
}

# Get Proxy Info
$bodyProxy = $null
If ($showProxy) {
    If ($null -ne $proxyList) {
        $arrProxy = $proxyList | Get-VBRProxyInfo | Select-Object @{Name = "Proxy Name"; Expression = { $_.ProxyName } },
        @{Name = "Transport Mode"; Expression = { $_.tMode } }, @{Name = "Max Tasks"; Expression = { $_.MaxTasks } },
        @{Name = "Proxy Host"; Expression = { $_.RealName } }, @{Name = "Host Type"; Expression = { $_.pType } },
        Enabled, @{Name = "IP Address"; Expression = { $_.IP } },
        @{Name = "RT (ms)"; Expression = { $_.Response } }, Status
        $bodyProxy = $arrProxy | Sort-Object "Proxy Host" |  ConvertTo-HTML -Fragment
        If ($arrProxy.Status -match "Dead") {
            $proxyHead = $subHead01err
        } ElseIf ($arrProxy -match "Alive") {
            $proxyHead = $subHead01suc
        } Else {
            $proxyHead = $subHead01
        }
        $bodyProxy = $proxyHead + "Proxy Details" + $subHead02 + $bodyProxy
    }
}

# Get Repository Info
$bodyRepo = $null
If ($showRepo) {
    If ($null -ne $repoList) {
        $arrRepo = $repoList | Get-VBRRepoInfo | Select-Object @{Name = "Repository Name"; Expression = { $_.Target } },
        @{Name = "Type"; Expression = { $_.rType } }, @{Name = "Max Tasks"; Expression = { $_.MaxTasks } },
        @{Name = "Host"; Expression = { $_.RepoHost } }, @{Name = "Path"; Expression = { $_.Storepath } },
        @{Name = "Free (GB)"; Expression = { $_.StorageFree } }, @{Name = "Total (GB)"; Expression = { $_.StorageTotal } },
        @{Name = "Free (%)"; Expression = { $_.FreePercentage } },
        @{Name = "Status"; Expression = {
                If ($_.FreePercentage -lt $repoCritical) { "Critical" }
                ElseIf ($_.StorageTotal -eq 0) { "Warning" }
                ElseIf ($_.FreePercentage -lt $repoWarn) { "Warning" }
                ElseIf ($_.FreePercentage -eq "Unknown") { "Unknown" }
                Else { "OK" } }
        }
        $bodyRepo = $arrRepo | Sort-Object "Repository Name" | ConvertTo-HTML -Fragment
        If ($arrRepo.status -match "Critical") {
            $repoHead = $subHead01err
        } ElseIf ($arrRepo.status -match "Warning|Unknown") {
            $repoHead = $subHead01war
        } ElseIf ($arrRepo.status -match "OK") {
            $repoHead = $subHead01suc
        } Else {
            $repoHead = $subHead01
        }
        $bodyRepo = $repoHead + "Repository Details" + $subHead02 + $bodyRepo
    }
}

# Get Scale Out Repository Info
$bodySORepo = $null
If ($showRepo) {
    If ($null -ne $repoListSo) {
        $arrSORepo = $repoListSo | Get-VBRSORepoInfo | Select-Object @{Name = "Scale Out Repository Name"; Expression = { $_.SOTarget } },
        @{Name = "Member Repository Name"; Expression = { $_.Target } }, @{Name = "Type"; Expression = { $_.rType } },
        @{Name = "Max Tasks"; Expression = { $_.MaxTasks } }, @{Name = "Host"; Expression = { $_.RepoHost } },
        @{Name = "Path"; Expression = { $_.Storepath } }, @{Name = "Free (GB)"; Expression = { $_.StorageFree } },
        @{Name = "Total (GB)"; Expression = { $_.StorageTotal } }, @{Name = "Free (%)"; Expression = { $_.FreePercentage } },
        @{Name = "Status"; Expression = {
                If ($_.FreePercentage -lt $repoCritical) { "Critical" }
                ElseIf ($_.StorageTotal -eq 0) { "Warning" }
                ElseIf ($_.FreePercentage -lt $repoWarn) { "Warning" }
                ElseIf ($_.FreePercentage -eq "Unknown") { "Unknown" }
                Else { "OK" } }
        }
        $bodySORepo = $arrSORepo | Sort-Object "Scale Out Repository Name", "Member Repository Name" | ConvertTo-HTML -Fragment
        If ($arrSORepo.status -match "Critical") {
            $sorepoHead = $subHead01err
        } ElseIf ($arrSORepo.status -match "Warning|Unknown") {
            $sorepoHead = $subHead01war
        } ElseIf ($arrSORepo.status -match "OK") {
            $sorepoHead = $subHead01suc
        } Else {
            $sorepoHead = $subHead01
        }
        $bodySORepo = $sorepoHead + "Scale Out Repository Details" + $subHead02 + $bodySORepo
    }
}

# Get Repository Agent User Permissions
$bodyRepoPerms = $null
If ($showRepoPerms) {
    If ($null -ne $repoList -or $null -ne $repoListSo) {
        $bodyRepoPerms = Get-RepoPermission | Select-Object Name, "Encryption Enabled", "Permission Type", Users | Sort-Object Name | ConvertTo-HTML -Fragment
        $bodyRepoPerms = $subHead01 + "Repository Permissions for Agent Jobs" + $subHead02 + $bodyRepoPerms
    }
}

# Get Replica Target Info
$bodyReplica = $null
If ($showReplicaTarget) {
    If ($null -ne $allJobsRp) {
        $repTargets = $allJobsRp | Get-VBRReplicaTarget | Select-Object @{Name = "Replica Target"; Expression = { $_.Target } }, Datastore,
        @{Name = "Free (GB)"; Expression = { $_.StorageFree } }, @{Name = "Total (GB)"; Expression = { $_.StorageTotal } },
        @{Name = "Free (%)"; Expression = { $_.FreePercentage } },
        @{Name = "Status"; Expression = {
                If ($_.FreePercentage -lt $replicaCritical) { "Critical" }
                ElseIf ($_.StorageTotal -eq 0) { "Warning" }
                ElseIf ($_.FreePercentage -lt $replicaWarn) { "Warning" }
                ElseIf ($_.FreePercentage -eq "Unknown") { "Unknown" }
                Else { "OK" }
            }
        } | Sort-Object "Replica Target"
        $bodyReplica = $repTargets | ConvertTo-HTML -Fragment
        If ($repTargets.status -match "Critical") {
            $reptarHead = $subHead01err
        } ElseIf ($repTargets.status -match "Warning|Unknown") {
            $reptarHead = $subHead01war
        } ElseIf ($repTargets.status -match "OK") {
            $reptarHead = $subHead01suc
        } Else {
            $reptarHead = $subHead01
        }
        $bodyReplica = $reptarHead + "Replica Target Details" + $subHead02 + $bodyReplica
    }
}

# Get Veeam Services Info
$bodyServices = $null
If ($showServices) {
    $vServers = Get-VeeamWinServer
    $vServices = Get-VeeamService $vServers
    If ($hideRunningSvc) { $vServices = $vServices | Where-Object { $_.Status -ne "Running" } }
    If ($null -ne $vServices) {
        $vServices = $vServices | Select-Object "Server Name", "Service Name",
        @{Name = "Status"; Expression = { If ($_.Status -eq "Stopped") { "Not Running" } Else { $_.Status } } }
        $bodyServices = $vServices | Sort-Object "Server Name", "Service Name" | ConvertTo-HTML -Fragment
        If ($vServices.status -match "Not Running") {
            $svcHead = $subHead01err
        } ElseIf ($vServices.status -notmatch "Running") {
            $svcHead = $subHead01war
        } ElseIf ($vServices.status -match "Running") {
            $svcHead = $subHead01suc
        } Else {
            $svcHead = $subHead01
        }
        $bodyServices = $svcHead + "Veeam Services (Windows)" + $subHead02 + $bodyServices
    }
}

# Get License Info
$bodyLicense = $null
If ($showLicExp) {
    $arrLicense = Get-VeeamSupportDate $vbrServer | Select-Object @{Name = "Expiry Date"; Expression = { $_.ExpDate } },
    @{Name = "Days Remaining"; Expression = { $_.DaysRemain } }, `
    @{Name = "Status"; Expression = {
            If ($_.DaysRemain -lt $licenseCritical) { "Critical" }
            ElseIf ($_.DaysRemain -lt $licenseWarn) { "Warning" }
            ElseIf ($_.DaysRemain -eq "Failed") { "Failed" }
            Else { "OK" } }
    }
    $bodyLicense = $arrLicense | ConvertTo-HTML -Fragment
    If ($arrLicense.Status -eq "OK") {
        $licHead = $subHead01suc
    } ElseIf ($arrLicense.Status -eq "Warning") {
        $licHead = $subHead01war
    } Else {
        $licHead = $subHead01err
    }
    $bodyLicense = $licHead + "License/Support Renewal Date" + $subHead02 + $bodyLicense
}

# Combine HTML Output
$htmlOutput = $headerObj + $bodyTop + $bodySummaryProtect + $bodySummaryBK + $bodySummaryRp + $bodySummaryBc + $bodySummaryTp + $bodySummaryEp + $bodySummarySb

If ($bodySummaryProtect + $bodySummaryBK + $bodySummaryRp + $bodySummaryBc + $bodySummaryTp + $bodySummaryEp + $bodySummarySb) {
    $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyMissing + $bodyWarning + $bodySuccess

If ($bodyMissing + $bodySuccess + $bodyWarning) {
    $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyMultiJobs

If ($bodyMultiJobs) {
    $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyJobsBk + $bodyJobSizeBk + $bodyAllSessBk + $bodyAllTasksBk + $bodyRunningBk + $bodyTasksRunningBk + $bodySessWFBk + $bodyTaskWFBk + $bodySessSuccBk + $bodyTaskSuccBk

If ($bodyJobsBk + $bodyJobSizeBk + $bodyAllSessBk + $bodyAllTasksBk + $bodyRunningBk + $bodyTasksRunningBk + $bodySessWFBk + $bodyTaskWFBk + $bodySessSuccBk + $bodyTaskSuccBk) {
    $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyRestoRunVM + $bodyRestoreVM

If ($bodyRestoRunVM + $bodyRestoreVM) {
    $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyJobsRp + $bodyAllSessRp + $bodyAllTasksRp + $bodyRunningRp + $bodyTasksRunningRp + $bodySessWFRp + $bodyTaskWFRp + $bodySessSuccRp + $bodyTaskSuccRp

If ($bodyJobsRp + $bodyAllSessRp + $bodyAllTasksRp + $bodyRunningRp + $bodyTasksRunningRp + $bodySessWFRp + $bodyTaskWFRp + $bodySessSuccRp + $bodyTaskSuccRp) {
    $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyJobsBc + $bodyJobSizeBc + $bodyAllSessBc + $bodyAllTasksBc + $bodySessIdleBc + $bodyTasksPendingBc + $bodyRunningBc + $bodyTasksRunningBc + $bodySessWFBc + $bodyTaskWFBc + $bodySessSuccBc + $bodyTaskSuccBc

If ($bodyJobsBc + $bodyJobSizeBc + $bodyAllSessBc + $bodyAllTasksBc + $bodySessIdleBc + $bodyTasksPendingBc + $bodyRunningBc + $bodyTasksRunningBc + $bodySessWFBc + $bodyTaskWFBc + $bodySessSuccBc + $bodyTaskSuccBc) {
    $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyJobsTp + $bodyAllSessTp + $bodyAllTasksTp + $bodyWaitingTp + $bodySessIdleTp + $bodyTasksPendingTp + $bodyRunningTp + $bodyTasksRunningTp + $bodySessWFTp + $bodyTaskWFTp + $bodySessSuccTp + $bodyTaskSuccTp

If ($bodyJobsTp + $bodyAllSessTp + $bodyAllTasksTp + $bodyWaitingTp + $bodySessIdleTp + $bodyTasksPendingTp + $bodyRunningTp + $bodyTasksRunningTp + $bodySessWFTp + $bodyTaskWFTp + $bodySessSuccTp + $bodyTaskSuccTp) {
    $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyTapes + $bodyTpPool + $bodyTpVlt + $bodyExpTp + $bodyTpExpPool + $bodyTpExpVlt + $bodyTpWrt

If ($bodyTapes + $bodyTpPool + $bodyTpVlt + $bodyExpTp + $bodyTpExpPool + $bodyTpExpVlt + $bodyTpWrt) {
    $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyJobsEp + $bodyJobSizeEp + $bodyAllSessEp + $bodyRunningEp + $bodySessWFEp + $bodySessSuccEp

If ($bodyJobsEp + $bodyJobSizeEp + $bodyAllSessEp + $bodyRunningEp + $bodySessWFEp + $bodySessSuccEp) {
    $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyJobsSb + $bodyAllSessSb + $bodyAllTasksSb + $bodyRunningSb + $bodyTasksRunningSb + $bodySessWFSb + $bodyTaskWFSb + $bodySessSuccSb + $bodyTaskSuccSb

If ($bodyJobsSb + $bodyAllSessSb + $bodyAllTasksSb + $bodyRunningSb + $bodyTasksRunningSb + $bodySessWFSb + $bodyTaskWFSb + $bodySessSuccSb + $bodyTaskSuccSb) {
    $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodySummaryConfig + $bodyProxy + $bodyRepo + $bodySORepo + $bodyRepoPerms + $bodyReplica + $bodyServices + $bodyLicense + $footerObj

# Fix Details
$htmlOutput = $htmlOutput.Replace("ZZbrZZ", "<br />")
# Remove trailing HTMLbreak
$htmlOutput = $htmlOutput.Replace("$($HTMLbreak + $footerObj)", "$($footerObj)")
# Add color to output depending on results

# Garantir que $htmlOutput esteja definido
if (-not $htmlOutput) {
    $htmlOutput = ""
}

# Substituir status com cores adequadas
$statusColors = @{
    'Running' = '#00b051'
    'OK' = '#00b051'
    'Alive' = '#00b051'
    'Success' = '#00b051'
    'Warning' = '#ffd96c'
    'Not Running' = '#ff6d6c'
    'Failed' = '#ff6d6c'
    'Critical' = '#ff6d6c'
    'Dead' = '#ff6d6c'
}

foreach ($status in $statusColors.Keys) {
    $color = $statusColors[$status]
    $htmlOutput = $htmlOutput -replace "<td>$status<", "<td style='color: $color;'>$status<"
}

# Verificar a presença de cada status antes de definir a cor do cabeçalho
$hasError = $htmlOutput -match '<td style="color: #ff6d6c;">'
$hasWarning = $htmlOutput -match '<td style="color: #ffd96c;">'
$hasSuccess = $htmlOutput -match '<td style="color: #00b051;">'

# Definir a cor do cabeçalho do relatório corretamente
if ($hasError) {
    $reportHeaderColor = '#ff6d6c'  # Vermelho apenas se houver erro
} elseif ($hasWarning) {
    $reportHeaderColor = '#ffd96c'  # Amarelo apenas se houver aviso
} elseif ($hasSuccess) {
    $reportHeaderColor = '#00b051'  # Verde apenas se houver sucesso
} else {
    $reportHeaderColor = '#00b051'  # Padrão para verde se não houver erros nem avisos
}

# Aplicar a cor do cabeçalho no relatório
$htmlOutput = $htmlOutput.Replace('ZZhdbgZZ', '#ffffff')

#endregion


# Save HTML Report to File
$htmlOutput | Out-File $pathHTML
#endregion

# Convert HTML to PDF

if (Test-Path $pathHTML) {

    # Executar o wkhtmltopdf para converter o HTML em PDF
    Start-Process -FilePath $wkhtmltopdfPath -ArgumentList "--enable-local-file-access", "`"$pathHTML`" `"$pathPDF`"" -NoNewWindow -Wait

    # Verificar se o PDF foi criado com sucesso
    if (Test-Path $pathPDF) {
        
        # Excluir o arquivo HTML após a conversão
        Remove-Item $pathHTML -Force
    } else {
        $licHead = $subHead01err
    }
} else {
    $licHead = $subHead01err
}