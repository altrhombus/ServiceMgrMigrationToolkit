# System Center Service Manager Migration Toolkit

## Synopsis
An automation toolkit to assist with migrating your Service Manager data to a new environment

## Description
This toolkit contains several PowerShell scripts to automate the process of exporting and importing the following Service Manager data:
* Incidents (including Activity Log and Work Item Attachments)
* Service Requests (including Activity Log, Work Item Attachments, and related Review/Manual/Parallel Activities) 

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
The current workflow separates the data export/import by Incidents and Service Requests.

## Notes
For your reference, here's some sites I found helpful:
* [http://www.zgc.se/index.php/2014/10/24/service-manager-powershell-examples/](http://www.zgc.se/index.php/2014/10/24/service-manager-powershell-examples/)
* [http://blogs.litware.se/?p=1369](http://blogs.litware.se/?p=1369)
* [https://blogs.technet.microsoft.com/servicemanager/2012/04/03/using-data-properties-from-the-parent-work-items-in-activity-email-templates/](https://blogs.technet.microsoft.com/servicemanager/2012/04/03/using-data-properties-from-the-parent-work-items-in-activity-email-templates/)
* [http://blogs.catapultsystems.com/mdowst/archive/2015/10/26/scsm-powershell-create-work-items/](http://blogs.catapultsystems.com/mdowst/archive/2015/10/26/scsm-powershell-create-work-items/)
* [https://gallery.technet.microsoft.com/System-Center-Server-7fddf821](https://gallery.technet.microsoft.com/System-Center-Server-7fddf821)
* [https://gallery.technet.microsoft.com/SCSM-Entity-Explorer-68b86bd2](https://gallery.technet.microsoft.com/SCSM-Entity-Explorer-68b86bd2)

Made with ❤️ in MKE
