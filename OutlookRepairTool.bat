@echo off
setlocal EnableDelayedExpansion

rem Set the path to the SCANPST.EXE tool
set SCANPST="C:\Program Files\Microsoft Office\root\Office16\SCANPST.EXE"

rem Set the directories to scan for Outlook files
set DIRS_TO_SCAN="C:\Users\xxxx\AppData\Local\Microsoft\Outlook"

rem Set the directory to store repair logs
set LOG_DIR="D:\OutlookRepairLog"

rem Set the directory to store backup files
set BACKUP_DIR="D:\OutlookBackup"

rem Display a message indicating the start of scanning Outlook files
echo Scanning Outlook files...

rem Loop through all PST and OST files in the specified directories
for /r "%DIRS_TO_SCAN%" %%a in (*.pst, *.ost) do (
  echo Repairing %%a ...
  
  rem Extract the base name of the file (without extension)
  set "file_basename=%%~na"
  
  rem Create a log file path using the base name and current date/time
  set "log_file=!LOG_DIR!\!file_basename!_!datetime!.log"
  
  rem Create a backup file path using the base name and current date/time
  set "backup_file=!BACKUP_DIR!\!file_basename!_!datetime!.bak"
  
  rem Run the SCANPST tool to repair the file, with specified options
  !SCANPST! -File "%%~fa" -Silent -Force -Rescan -BackupFile "!backup_file!" -LogFile "!log_file!" -Telemetry -CorruptionLog "!log_file!.corruption.log"
  
  rem Append the current date/time and file path to the log file
  echo !datetime! - %%a >> "!log_file!"
)

rem Display a message indicating the start of moving backup files
echo Moving backup files...

rem Loop through all backup files in the specified directories
for /r "%DIRS_TO_SCAN%" %%a in (*.pst.bak, *.ost.bak) do (
  set "file_basename=%%~na"
  
  rem Create a backup file path using the base name and current date/time
  set "backup_file=!BACKUP_DIR!\!file_basename!_!datetime!.bak"
  
  rem Check if a backup file with the same name already exists
  if exist "!backup_file!" (
    rem If it exists, add a counter to the file name
    set "count=1"
    set "backup_file=!BACKUP_DIR!\!file_basename!_!datetime!_!count!.bak"
    
    rem Loop until a unique file name is found
    :loop
    if exist "!backup_file!" (
      set /a "count+=1"
      set "backup_file=!BACKUP_DIR!\!file_basename!_!datetime!_!count!.bak"
      goto :loop
    )
  )
  
  rem Move the backup file to the specified backup directory
  move "%%~fa" "!backup_file!" >nul
)

rem Display a message indicating the start of moving log files
echo Moving log files...

rem Loop through all log files in the specified log directory
for /r "%LOG_DIR%" %%a in (*.log, *.corruption.log) do (
  set "file_basename=%%~na"
  
  rem Create a log file path using the base name and current date/time
  set "log_file=!LOG_DIR!\!file_basename!_!datetime!.log"
  
  rem Check if a log file with the same name already exists
  if exist "!log_file!" (
    rem If it exists, add a counter to the file name
    set "count=1"
    set "log_file=!LOG_DIR!\!file_basename!_!datetime!_!count!.log"
    
    rem Loop until a unique file name is found
    :loop
    if exist "!log_file!" (
      set /a "count+=1"
      set "log_file=!LOG_DIR!\!file_basename!_!datetime!_!count!.log"
      goto :loop
    )
  )
  
  rem Move the log file to the specified log directory
  move "%%~fa" "!log_file!" >nul
)

rem Display a message indicating the completion of file repairs
echo All files repaired.

rem Pause the script execution to keep the command prompt window open
pause
