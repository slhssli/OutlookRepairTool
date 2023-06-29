# Outlook Repair Tool

The Outlook Repair Tool is a batch script that automates the scanning, repair, and backup of Outlook PST and OST files using the Microsoft SCANPST.EXE utility. It also provides functionality to move backup and log files to designated directories.
Prerequisites

    Microsoft Office installed (including SCANPST.EXE)
    Windows operating system

Usage

    Edit the batch script OutlookRepairTool.bat and modify the following variables according to your environment:
        SCANPST: Path to the SCANPST.EXE tool.
        DIRS_TO_SCAN: Directories to scan for Outlook files (PST and OST).
        LOG_DIR: Directory to store repair log files.
        BACKUP_DIR: Directory to store backup files.

    Save the modified batch script.

    Double-click the OutlookRepairTool.bat file to execute the script.

    The script will scan the specified directories for PST and OST files, repair them using SCANPST.EXE, and generate repair log files.

    Backup files will be created in the designated backup directory.

    The script will move the backup and log files to their respective directories.

    Once the process is complete, a message will be displayed indicating that all files have been repaired.

    You can press any key to exit the script.

Notes

    It is recommended to create a backup of your Outlook files before running this script.
    The script may take some time to complete depending on the number and size of the Outlook files being scanned.
    Make sure to run the script with administrative privileges if necessary.

License

This script is provided under the MIT License.

Feel free to modify and adapt the script according to your needs.
Disclaimer

    The Outlook Repair Tool script is provided as-is without any warranty. Use it at your own risk.
    Make sure to test the script on a small subset of files before running it on your entire Outlook data.

Credits

This script was created by Ali Alhamed.

If you have any questions or encounter any issues, please contact at https://alialhamed.com/contact-us/
