# InstallerCleaner

Following what you do with your PC, the C:/Windows/Installer directory can reach several Gb of data to save unnecessary MSI and MSP files, normally used for installation and removal of the applications and software patches. These files remain in the C:/Windows/Installer directory even if you delete/uninstall your software in a clean manner.
Several tools propose to clean the C:/Windows/Installer directory. But even if they are efficient, I was not happy with the usage. They request too much clicks or don't have backup options.

InstallerCleaner is a small tool to clean the C:/Windows/Installer directory and offers backup capability. It is a command-line tool that requests admin rights. It can generate a report or perform simulation of files removal. So, you know what will be done or what was done by the tool. If you run the software with the appropriate options, you can easily recover your deleted files.

## Usage
The tool has three main options:
- simulate: to identify the files to remove
    command-line --simulate path_to_report_file
- delete: to delete the unnecessary files
- move: to move the unnecessary files to an identified folder (-backup mode)
    command-line --move target_directory

Additionally, the tool can create a report (the report is automatically generated in simulate mode), thanks to the command line --report.
The command-line options are also accessible thanks the --help command.

## Installation
Download the files set in bin/Release, copy somewhere on your hard drive, it could be C:/Temp... and launch the tools from a Windows Command box.

## Build
As there is a single module, it shall not be a problem.
The build was done with .Net Framework 4.8, but you can change it without a problem, as it does not use the most advanced functionalities.
The command-line interface is built with the CommandParser dll; you can retrieve the code from https://github.com/NagoDede/CommandParser. The dll is available in the bin directory.

## Additional information
The identification of the critical MSI / MSP files (files not deleted) is made thanks to WindowsInstaller, which retrieve the data in the registry table.

Even if slightly outdated, you can read the page https://www.raymond.cc/blog/safely-delete-unused-msi-and-mst-files-from-windows-installer-folder/ to learn more about the /Installer directory.
