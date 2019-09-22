# O365-Automated-Installer

These batch files were included with a landesk client installer (Default Windows Configuration.exe)

The goal was to make a tool that could be loaded on a flash drive and install office365 automatically. autorun.inf was not an option because the autorun feature was disabled.

The work was eventually turned into making a tool that a tech with little to no training and limited knowledge of English could deploy to Windows desktops internationally. Hurdles included supporting localization with languages that use non latin characters or differently named folders, being able to support multiple operating systems, running on both 64 or 32 bit systems, and detecting then uninstalling any previous installed copies of office to avoid installation issues or follow up requests.

# Usage

This installer assumes that the exact directory structure is sitting on the root of a usb flash drive

Run the InstallOffice.bat for your architecture (64/32 bit). 

The script will do the following:

1) Find the drive letter for the usb drive that the installer is on without prompting the user
2) Detect operating system
3) If OS is found to be windows 7 or earlier a restart flag will be triggered for the uninstall of office to be successful, and a resume script will be added to the flash drive (setting scheduled tasks was not an option for users running this, this part may be changed to your taste)
4) Exits any running office applications
5) Runs the vbs uninstallers gathered from the microsoft fixit tool for all older versions of office
6) Installs office and Teams
7) Per manager request, copies data to the landesk folder
8) Pauses to show progress has completed without errors, then exits

# Troubleshooting

There is a troubleshooting folder for installation issues. At the time that this tool was written the biggest issues plaguing the enterprise was windows not showing as activated (WindowsActivation.bat), cached office credentials interfering with user sign in on the new version of office (Delete Credential Vault.bat), or refusing to pull a license for office (O365 Activation fix.bat)

In addition because language support was needed, if the language could not be found automatically then language packs are made available in the Languages folder.
