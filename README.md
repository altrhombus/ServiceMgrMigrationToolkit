# System Center Service Manager Migration Toolkit

## Synopsis
An automation toolkit to assist with migrating your Service Manager data to a new environment

## Description
This toolkit contains several PowerShell scripts to automate the process of exporting and importing the following Service Manager data:
* Incidents (including Activity Log and Work Item Attachments)
* Service Requests (including Activity Log, Work Item Attachments, and related Review/Manual Activities) 

### Prerequisites
#### Dependencies
This toolkit requires the following:
* [SMLets PowerShell Cmdlets](http://smlets.codeplex.com/)
* [Windows Management Framework 4.0 or greater](https://msdn.microsoft.com/en-us/powershell/wmf/readme)

#### Tested Environments
This script has been tested against 2012 R2 (exporting) and 2016 (importing). If you have success with other environments, let me know (including URs)!

#### Management Packs and Enumerations
In it's current form, the toolkit does not migrate your Management Packs. When importing in the new environment, any enum values are matched by name (for now). To prevent import issues, you should do the following:
* Import any customized Management Packs to your new environment
* Set up your AD connector and make sure it completes a sync
Some basic sanity checking is performed prior to the import to verify that your lists are populated correctly.

## Example
The current workflow separates the data export/import by Incidents, Service Requests, and Configuration Items (in particular, the Cireson Asset Management data).

## Notes
For your reference, here's some sites I found helpful:
< ... >

Made with ❤️ in MKE
