# These PowerShell scripts are currently in BETA. Please test each script individually and edit as necessary before deploying to production.

# Microsoft 365 PowerShell Scripts

A collection of PowerShell scripts for managing, auditing, and reporting on Microsoft 365 tenants.

Each script includes built-in documentation, parameter support, and clear output.

Scripts are provided as-is. Review and test before running in production.

These scripts have been tested in my HomeLab environment using Windows Server, Virtual Machines, Windows Sandbox, AutomatedLab, and Pester. 

They have **not** been tested in a production or enterprise environment such as Microsoft 365 or Office 365.

# Acknowledgements

This repository has benefited from thoughtful community feedback. Special thanks to:

Snickasaurus (https://github.com/Snickasaurus) —

Provided multiple suggestions that directly improved the repository, including:

Recommending the creation of 00-Install-365Modules.ps1 to consolidate module installation across all scripts

Advising the prepending of zeros to scripts 1–9 for proper sorting

Highlighting that 12-Get-M365RiskySignInsReport.ps1 is a template and suggesting clarification

Noting formatting issues (extra slashes) in scripts 42–50

Thanks to these contributions, the repository is now cleaner, better organized, and easier to use for administrators managing Microsoft 365 environments.
