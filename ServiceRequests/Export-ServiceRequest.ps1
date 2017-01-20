<#
.SYNOPSIS
    Exports Service Requests from System Center Service Manager to a CSV for simplified importing to other environments

.DESCRIPTION
    Exports all Service Requests and their useful/related data to a CSV file, which can be used for good or for awesome

    The purpose of this script is to help automate the process of migrating to a new Service Manager environment, bringing
    only the data you want. In this script's current life, we grab all service requests that can be found using Get-SCSMObject and
    write out several files:
        - A CSV containing all service requests, their data, and their assigned to/affected by related users
        - A CSV containing all activity log entries for all service requests
        - A CSV containing all manual activities related to service requests
        - A CSV containing all review activities related to service requests
        - A CSV containing all parallel activities related to service requests
    With these CSVs, you can re-import into a new environment, or quickly populate a development environment

.PARAMETER csvOutput
    The path and name of a CSV file that is created containing all Service Request data
    
    This field is required, and should be a location that you have write access to

.PARAMETER csvOutputActivityLog
    The path and name of a CSV file that is created containing all Service Request Activity Log data

    This field is required, and should be a location that you have write access to

.PARAMETER csvOutputRelatedManualActivities
    The path and name of a CSV file that is created containing all related Manual Activities data

    This field is required, and should be a location that you have write access to

.PARAMETER csvOutputRelatedReviewActivities
    The path and name of a CSV file that is created containing all related Review Activities data

    This field is required, and should be a location that you have write access to

.PARAMETER csvOutputRelatedParallelActivities
    The path and name of a CSV file that is created containing all related Parallel Activities data

    This field is required, and should be a location that you have write access to

.EXAMPLE
    Export all Service Requests from a Service Manager environment

    Export-ServiceRequest.ps1 -csvOutput .\servicerequests.csv -csvOutputActivityLog .\sr_al_output.csv -csvOutputRelatedManualActivities .\sr_ma_output.csv -csvOutputRelatedReviewActivities .\sr_ra_output.csv -csvOutputRelatedParallelActivities .\sr_pa_output.csv

.NOTES
    Requires the Windows Management Framework 4 or greater
    Requires the SMlets PowerShell module
    Tested with System Center Service Manager 2012 R2 and 2016

    Jacob Thornberry
    @jakertberry
    Made with ❤ in MKE
#>

[CmdletBinding()]
Param (
    [Parameter(Mandatory = $True, HelpMessage = "Where should we write out the list of Service Requests in CSV format")]
    $csvOutput,

    [Parameter(Mandatory = $True, HelpMessage = "Where should we write out the SR Activity Log CSV")]
    $csvOutputActivityLog,

    [Parameter(Mandatory = $True, HelpMessage = "Where should we write out the related SR manual activities")]
    $csvOutputRelatedManualActivities,
    
    [Parameter(Mandatory = $True, HelpMessage = "Where should we write out the related SR review activities")]
    $csvOutputRelatedReviewActivities,

    [Parameter(Mandatory = $True, HelpMessage = "Where should we write out the related SR parallel activities")]
    $csvOutputRelatedParallelActivities
)

function Test-OutputPath([string]$filePath, [string]$serviceType)
{
    if (!(Test-Path $filePath))
    {
        Write-Error "The path $filePath does not exist for $serviceType" -ErrorAction Stop
    }
}

function Import-SMLets
{
    Write-Progress -Activity "Initializing" -Status "Checking for SMLets"
    if (!(Get-Module SMLets))
    {
        Write-Progress -Activity "Initializing" -Status "Importing the SMLets Module"
        Import-Module SMLets -Force -ErrorAction Stop   
    }
}

function Export-ServiceRequest([string]$filePath)
{
    Write-Progress -Activity "Initializing" -Status "Reticulating Splines"
    $i = 0
    $id = 0
    $class = Get-SCSMClass -Name System.WorkItem.ServiceRequest$
    $userClass = Get-SCSMClass -Name System.Domain.User$  
    $affectedUserRelClass = Get-SCSMRelationshipClass -Name System.WorkItemAffectedUser$
    $assignedToUserRelClass  = Get-SCSMRelationshipClass -Name System.WorkItemAssignedToUser$ 

    $srArray =@()

    $totalCount = (Get-SCSMObject -Class $class).Count

    Get-SCSMObject -Class $class | ForEach-Object {
        $id = $_.Id
        Write-Progress -Activity "Exporting Service Requests" -Status "Found $id which affects $affectedUser and was assigned to $assignedTo" -PercentComplete (($i / $totalCount) * 100)
        $srObject = New-Object PSObject
        
        $affectedUser = Get-SCSMRelatedObject -SMObject $_ -Relationship $affectedUserRelClass
        $assignedTo = Get-SCSMRelatedObject -SMObject $_ -Relationship $assignedToUserRelClass
        
        Write-Verbose "$id affects $affectedUser and has been assigned to $assignedTo"

        $srObject | Add-Member -MemberType NoteProperty -Name "Id" -Value $_.Id
        $srObject | Add-Member -MemberType NoteProperty -Name "Title" -Value $_.Title
        $srObject | Add-Member -MemberType NoteProperty -Name "Description" -Value $_.Description
        $srObject | Add-Member -MemberType NoteProperty -Name "Status" -Value $_.Status
        $srObject | Add-Member -MemberType NoteProperty -Name "StatusDisplayName" -Value $_.Status.DisplayName
        $srObject | Add-Member -MemberType NoteProperty -Name "TemplateId" -Value $_.TemplateId
        $srObject | Add-Member -MemberType NoteProperty -Name "Priority" -Value $_.Priority
        $srObject | Add-Member -MemberType NoteProperty -Name "PriorityDisplayName" -Value $_.Priority.DisplayName
        $srObject | Add-Member -MemberType NoteProperty -Name "Urgency" -Value $_.Urgency
        $srObject | Add-Member -MemberType NoteProperty -Name "UrgencyDisplayName" -Value $_.Urgency.DisplayName
        $srObject | Add-Member -MemberType NoteProperty -Name "CompletedDate" -Value $_.CompletedDate
        $srObject | Add-Member -MemberType NoteProperty -Name "ClosedDate" -Value $_.ClosedDate
        $srObject | Add-Member -MemberType NoteProperty -Name "Source" -Value $_.Source
        $srObject | Add-Member -MemberType NoteProperty -Name "SourceDisplayName" -Value $_.Source.DisplayName
        $srObject | Add-Member -MemberType NoteProperty -Name "ImplementationResults" -Value $_.ImplementationResults
        $srObject | Add-Member -MemberType NoteProperty -Name "ImplementationResultsDisplayName" -Value $_.ImplementationResults.DisplayName
        $srObject | Add-Member -MemberType NoteProperty -Name "Notes" -Value $_.Notes
        $srObject | Add-Member -MemberType NoteProperty -Name "Area" -Value $_.Area
        $srObject | Add-Member -MemberType NoteProperty -Name "AreaDisplayName" -Value $_.Area.DisplayName
        $srObject | Add-Member -MemberType NoteProperty -Name "SupportGroup" -Value $_.SupportGroup
        $srObject | Add-Member -MemberType NoteProperty -Name "SupportGroupDisplayName" -Value $_.SupportGroup.DisplayName
        $srObject | Add-Member -MemberType NoteProperty -Name "ContactMethod" -Value $_.ContactMethod
        $srObject | Add-Member -MemberType NoteProperty -Name "CreatedDate" -Value $_.CreatedDate
        $srObject | Add-Member -MemberType NoteProperty -Name "ScheduledStartDate" -Value $_.ScheduledStartDate
        $srObject | Add-Member -MemberType NoteProperty -Name "ScheduledEndDate" -Value $_.ScheduledEndDate
        $srObject | Add-Member -MemberType NoteProperty -Name "ActualStartDate" -Value $_.ActualStartDate
        $srObject | Add-Member -MemberType NoteProperty -Name "ActualEndDate" -Value $_.ActualEndDate
        $srObject | Add-Member -MemberType NoteProperty -Name "IsDowntime" -Value ("$" + $_.IsDowntime)
        $srObject | Add-Member -MemberType NoteProperty -Name "IsParent" -Value ("$" + $_.IsParent)
        $srObject | Add-Member -MemberType NoteProperty -Name "ScheduledDowntimeStartDate" -Value $_.ScheduledDowntimeStartDate
        $srObject | Add-Member -MemberType NoteProperty -Name "ScheduledDowntimeEndDate" -Value $_.ScheduledDowntimeEndDate
        $srObject | Add-Member -MemberType NoteProperty -Name "ActualDowntimeStartDate" -Value $_.ActualDowntimeStartDate
        $srObject | Add-Member -MemberType NoteProperty -Name "ActualDowntimeEndDate" -Value $_.ActualDowntimeEndDate
        $srObject | Add-Member -MemberType NoteProperty -Name "RequiredBy" -Value $_.RequiredBy
        $srObject | Add-Member -MemberType NoteProperty -Name "PlannedCost" -Value $_.PlannedCost
        $srObject | Add-Member -MemberType NoteProperty -Name "ActualCost" -Value $_.ActualCost
        $srObject | Add-Member -MemberType NoteProperty -Name "PlannedWork" -Value $_.PlannedWork
        $srObject | Add-Member -MemberType NoteProperty -Name "ActualWork" -Value $_.ActualWork
        $srObject | Add-Member -MemberType NoteProperty -Name "UserInput" -Value $_.UserInput
        $srObject | Add-Member -MemberType NoteProperty -Name "FirstAssignedDate" -Value $_.FirstAssignedDate
        $srObject | Add-Member -MemberType NoteProperty -Name "FirstResponseDate" -Value $_.FirstResponseDate
        $srObject | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $_.DisplayName
        $srObject | Add-Member -MemberType NoteProperty -Name "Name" -Value $_.Name
        $srObject | Add-Member -MemberType NoteProperty -Name "Path" -Value $_.Path
        $srObject | Add-Member -MemberType NoteProperty -Name "FullName" -Value $_.FullName
        $srObject | Add-Member -MemberType NoteProperty -Name "TimeAdded" -Value $_.TimeAdded
        $srObject | Add-Member -MemberType NoteProperty -Name "LastModifiedBy" -Value $_.LastModifiedBy #GET ENUM
        $srObject | Add-Member -MemberType NoteProperty -Name "LastModified" -Value $_.LastModified
        $srObject | Add-Member -MemberType NoteProperty -Name "IsNew" -Value ("$" + $_.IsNew)
        $srObject | Add-Member -MemberType NoteProperty -Name "HasChanges" -Value ("$" + $_.HasChanges)
        $srObject | Add-Member -MemberType NoteProperty -Name "GroupsAsDifferentType" -Value ("$" + $_.GroupsAsDifferentType)
        $srObject | Add-Member -MemberType NoteProperty -Name "ViewName" -Value $_.ViewName
        $srObject | Add-Member -MemberType NoteProperty -Name "ObjectMode" -Value $_.ObjectMode
        $srObject | Add-Member -MemberType NoteProperty -Name "AffectedUser" -Value $affectedUser
        $srObject | Add-Member -MemberType NoteProperty -Name "AssignedTo" -Value $assignedTo
        
        $srArray += $srObject
        $i++
    }

    $srArray | Export-Csv $filePath
    
    $srObject = $null
    $srArray = $null
}

function Export-RelatedManualActivity([string]$filePath)
{
    $i = 0
    $class = Get-SCSMClass -Name System.WorkItem.ServiceRequest$
    $userClass = Get-SCSMClass -Name System.Domain.User$

    $srArray =@()

    $totalCount = (Get-SCSMObject -Class $class).Count

    Get-SCSMObject -Class $class | ForEach-Object {
        $id = $_.Id
        $maCount = (Get-SCSMRelatedObject -SMObject $_ | Where-Object {$_.ClassName -eq "System.WorkItem.Activity.ManualActivity"}).Count
        Write-Verbose "$id has $maCount Manual Activities"
        Write-Progress -Activity "Exporting Related Manual Activities" -Status "Exported $maCount Manual Activities which were related to $id" -PercentComplete (($i / $totalCount) * 100)
        Get-SCSMRelatedObject -SMObject $_ | Where-Object {$_.ClassName -eq "System.WorkItem.Activity.ManualActivity"} | ForEach-Object {
            
            $srObject = New-Object PSObject
            $srObject | Add-Member -MemberType NoteProperty -Name "Parent" -Value $id
            $srObject | Add-Member -MemberType NoteProperty -Name "SequenceId" -Value $_.SequenceId
            $srObject | Add-Member -MemberType NoteProperty -Name "ChildId" -Value $_.ChildId
            $srObject | Add-Member -MemberType NoteProperty -Name "Notes" -Value $_.Notes
            $srObject | Add-Member -MemberType NoteProperty -Name "Status" -Value $_.Status
            $srObject | Add-Member -MemberType NoteProperty -Name "StatusDisplayName" -Value $_.Status.DisplayName
            $srObject | Add-Member -MemberType NoteProperty -Name "Priority" -Value $_.Priority
            $srObject | Add-Member -MemberType NoteProperty -Name "PriorityDisplayName" -Value $_.Priority.DisplayName
            $srObject | Add-Member -MemberType NoteProperty -Name "Area" -Value $_.Area
            $srObject | Add-Member -MemberType NoteProperty -Name "AreaDisplayName" -Value $_.Area.DisplayName
            $srObject | Add-Member -MemberType NoteProperty -Name "Stage" -Value $_.Stage
            $srObject | Add-Member -MemberType NoteProperty -Name "StageDisplayName" -Value $_.Stage.DisplayName
            $srObject | Add-Member -MemberType NoteProperty -Name "Documentation" -Value $_.Documentation
            $srObject | Add-Member -MemberType NoteProperty -Name "Skip" -Value ("$" + $_.Skip)
            $srObject | Add-Member -MemberType NoteProperty -Name "Id" -Value $_.Id
            $srObject | Add-Member -MemberType NoteProperty -Name "Title" -Value $_.Title
            $srObject | Add-Member -MemberType NoteProperty -Name "Description" -Value $_.Description
            $srObject | Add-Member -MemberType NoteProperty -Name "ContactMethod" -Value $_.ContactMethod
            $srObject | Add-Member -MemberType NoteProperty -Name "CreatedDate" -Value $_.CreatedDate
            $srObject | Add-Member -MemberType NoteProperty -Name "ScheduledStartDate" -Value $_.ScheduledStartDate
            $srObject | Add-Member -MemberType NoteProperty -Name "ScheduledEndDate" -Value $_.ScheduledEndDate
            $srObject | Add-Member -MemberType NoteProperty -Name "ActualStartDate" -Value $_.ActualStartDate
            $srObject | Add-Member -MemberType NoteProperty -Name "ActualEndDate" -Value $_.ActualEndDate
            $srObject | Add-Member -MemberType NoteProperty -Name "IsDowntime" -Value ("$" + $_.IsDowntime)
            $srObject | Add-Member -MemberType NoteProperty -Name "IsParent" -Value ("$" + $_.IsParent)
            $srObject | Add-Member -MemberType NoteProperty -Name "ScheduledDowntimeStartDate" -Value $_.ScheduledDowntimeStartDate
            $srObject | Add-Member -MemberType NoteProperty -Name "ScheduledDowntimeEndDate" -Value $_.ScheduledDowntimeEndDate
            $srObject | Add-Member -MemberType NoteProperty -Name "ActualDowntimeStartDate" -Value $_.ActualDowntimeStartDate
            $srObject | Add-Member -MemberType NoteProperty -Name "ActualDowntimeEndDate" -Value $_.ActualDowntimeEndDate
            $srObject | Add-Member -MemberType NoteProperty -Name "RequiredBy" -Value $_.RequiredBy
            $srObject | Add-Member -MemberType NoteProperty -Name "PlannedCost" -Value $_.PlannedCost
            $srObject | Add-Member -MemberType NoteProperty -Name "ActualCost" -Value $_.ActualCost
            $srObject | Add-Member -MemberType NoteProperty -Name "PlannedWork" -Value $_.PlannedWork
            $srObject | Add-Member -MemberType NoteProperty -Name "ActualWork" -Value $_.ActualWork
            $srObject | Add-Member -MemberType NoteProperty -Name "UserInput" -Value $_.UserInput
            $srObject | Add-Member -MemberType NoteProperty -Name "FirstAssignedDate" -Value $_.FirstAssignedDate
            $srObject | Add-Member -MemberType NoteProperty -Name "FirstResponseDate" -Value $_.FirstResponseDate
            $srObject | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $_.DisplayName
            $srObject | Add-Member -MemberType NoteProperty -Name "Name" -Value $_.Name
            $srObject | Add-Member -MemberType NoteProperty -Name "Path" -Value $_.Path
            $srObject | Add-Member -MemberType NoteProperty -Name "FullName" -Value $_.FullName
            $srObject | Add-Member -MemberType NoteProperty -Name "TimeAdded" -Value $_.TimeAdded
            $srObject | Add-Member -MemberType NoteProperty -Name "LastModifiedBy" -Value $_.LastModifiedBy
            $srObject | Add-Member -MemberType NoteProperty -Name "LastModified" -Value $_.LastModified
            $srObject | Add-Member -MemberType NoteProperty -Name "IsNew" -Value ("$" + $_.IsNew)
            $srObject | Add-Member -MemberType NoteProperty -Name "HasChanges" -Value ("$" + $_.HasChanges)
            $srObject | Add-Member -MemberType NoteProperty -Name "GroupsAsDifferentType" -Value ("$" + $_.GroupsAsDifferentType)
            $srObject | Add-Member -MemberType NoteProperty -Name "ViewName" -Value $_.ViewName
            $srObject | Add-Member -MemberType NoteProperty -Name "ObjectMode" -Value $_.ObjectMode
            $srArray += $srObject
        }
        
        $i++
    }

    $srArray | Export-Csv $filePath
    $srObject = $null
    $srArray = $null
}

function Export-RelatedManualActivityInsideAParallelActivity([string]$filePath)
# I'm really not happy with this code. The opportunity for optimization here is immense.
# Also, it only checks one level deep, and only checks for manual activities.
# For now, It Gets the Job Done (TM)
{
    $i = 0
    $class = Get-SCSMClass -Name System.WorkItem.ServiceRequest$
    $userClass = Get-SCSMClass -Name System.Domain.User$

    $srArray =@()

    $totalCount = (Get-SCSMObject -Class $class).Count

    Get-SCSMObject -Class $class | ForEach-Object {
        $id = $_.Id
        $paCount = (Get-SCSMRelatedObject -SMObject $_ | Where-Object {$_.ClassName -eq "System.WorkItem.Activity.ParallelActivity"}).Count
        Write-Verbose "$id has $paCount Parallel Activities"
        Write-Progress -Activity "Exporting Parallel Activities and their related Manual Activities" -Status "Processing $id" -PercentComplete (($i / $totalCount) * 100)

        Get-SCSMRelatedObject -SMObject $_ | Where-Object {$_.ClassName -eq "System.WorkItem.Activity.ParallelActivity"} | ForEach-Object {
            $srObject = New-Object PSObject
            $srObject | Add-Member -MemberType NoteProperty -Name "Parent" -Value $id
            $srObject | Add-Member -MemberType NoteProperty -Name "ClassName" -Value "System.WorkItem.Activity.ParallelActivity"
            $srObject | Add-Member -MemberType NoteProperty -Name "SequenceId" -Value $_.SequenceId
            $srObject | Add-Member -MemberType NoteProperty -Name "ChildId" -Value $_.ChildId
            $srObject | Add-Member -MemberType NoteProperty -Name "Notes" -Value $_.Notes
            $srObject | Add-Member -MemberType NoteProperty -Name "Status" -Value $_.Status
            $srObject | Add-Member -MemberType NoteProperty -Name "StatusDisplayName" -Value $_.Status.DisplayName
            $srObject | Add-Member -MemberType NoteProperty -Name "Priority" -Value $_.Priority
            $srObject | Add-Member -MemberType NoteProperty -Name "PriorityDisplayName" -Value $_.Priority.DisplayName
            $srObject | Add-Member -MemberType NoteProperty -Name "Area" -Value $_.Area
            $srObject | Add-Member -MemberType NoteProperty -Name "AreaDisplayName" -Value $_.Area.DisplayName
            $srObject | Add-Member -MemberType NoteProperty -Name "Stage" -Value $_.Stage
            $srObject | Add-Member -MemberType NoteProperty -Name "StageDisplayName" -Value $_.Stage.DisplayName
            $srObject | Add-Member -MemberType NoteProperty -Name "Documentation" -Value $_.Documentation
            $srObject | Add-Member -MemberType NoteProperty -Name "Skip" -Value ("$" + $_.Skip)
            $srObject | Add-Member -MemberType NoteProperty -Name "Id" -Value $_.Id
            $parentParallelActivity = $_.Id
            $srObject | Add-Member -MemberType NoteProperty -Name "Title" -Value $_.Title
            $srObject | Add-Member -MemberType NoteProperty -Name "Description" -Value $_.Description
            $srObject | Add-Member -MemberType NoteProperty -Name "ContactMethod" -Value $_.ContactMethod
            $srObject | Add-Member -MemberType NoteProperty -Name "CreatedDate" -Value $_.CreatedDate
            $srObject | Add-Member -MemberType NoteProperty -Name "ScheduledStartDate" -Value $_.ScheduledStartDate
            $srObject | Add-Member -MemberType NoteProperty -Name "ScheduledEndDate" -Value $_.ScheduledEndDate
            $srObject | Add-Member -MemberType NoteProperty -Name "ActualStartDate" -Value $_.ActualStartDate
            $srObject | Add-Member -MemberType NoteProperty -Name "ActualEndDate" -Value $_.ActualEndDate
            $srObject | Add-Member -MemberType NoteProperty -Name "IsDowntime" -Value ("$" + $_.IsDowntime)
            $srObject | Add-Member -MemberType NoteProperty -Name "IsParent" -Value ("$" + $_.IsParent)
            $srObject | Add-Member -MemberType NoteProperty -Name "ScheduledDowntimeStartDate" -Value $_.ScheduledDowntimeStartDate
            $srObject | Add-Member -MemberType NoteProperty -Name "ScheduledDowntimeEndDate" -Value $_.ScheduledDowntimeEndDate
            $srObject | Add-Member -MemberType NoteProperty -Name "ActualDowntimeStartDate" -Value $_.ActualDowntimeStartDate
            $srObject | Add-Member -MemberType NoteProperty -Name "ActualDowntimeEndDate" -Value $_.ActualDowntimeEndDate
            $srObject | Add-Member -MemberType NoteProperty -Name "RequiredBy" -Value $_.RequiredBy
            $srObject | Add-Member -MemberType NoteProperty -Name "PlannedCost" -Value $_.PlannedCost
            $srObject | Add-Member -MemberType NoteProperty -Name "ActualCost" -Value $_.ActualCost
            $srObject | Add-Member -MemberType NoteProperty -Name "PlannedWork" -Value $_.PlannedWork
            $srObject | Add-Member -MemberType NoteProperty -Name "ActualWork" -Value $_.ActualWork
            $srObject | Add-Member -MemberType NoteProperty -Name "UserInput" -Value $_.UserInput
            $srObject | Add-Member -MemberType NoteProperty -Name "FirstAssignedDate" -Value $_.FirstAssignedDate
            $srObject | Add-Member -MemberType NoteProperty -Name "FirstResponseDate" -Value $_.FirstResponseDate
            $srObject | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $_.DisplayName
            $srObject | Add-Member -MemberType NoteProperty -Name "Name" -Value $_.Name
            $srObject | Add-Member -MemberType NoteProperty -Name "Path" -Value $_.Path
            $srObject | Add-Member -MemberType NoteProperty -Name "FullName" -Value $_.FullName
            $srObject | Add-Member -MemberType NoteProperty -Name "TimeAdded" -Value $_.TimeAdded
            $srObject | Add-Member -MemberType NoteProperty -Name "LastModifiedBy" -Value $_.LastModifiedBy
            $srObject | Add-Member -MemberType NoteProperty -Name "LastModified" -Value $_.LastModified
            $srObject | Add-Member -MemberType NoteProperty -Name "IsNew" -Value ("$" + $_.IsNew)
            $srObject | Add-Member -MemberType NoteProperty -Name "HasChanges" -Value ("$" + $_.HasChanges)
            $srObject | Add-Member -MemberType NoteProperty -Name "GroupsAsDifferentType" -Value ("$" + $_.GroupsAsDifferentType)
            $srObject | Add-Member -MemberType NoteProperty -Name "ViewName" -Value $_.ViewName
            $srObject | Add-Member -MemberType NoteProperty -Name "ObjectMode" -Value $_.ObjectMode
            $srArray += $srObject

            Get-SCSMRelatedObject -SMObject $_ | Where-Object {$_.ClassName -eq "System.WorkItem.Activity.ManualActivity"} | ForEach-Object {
                $srObject = New-Object PSObject
                $srObject | Add-Member -MemberType NoteProperty -Name "Parent" -Value $parentParallelActivity
                $srObject | Add-Member -MemberType NoteProperty -Name "ClassName" -Value "System.WorkItem.Activity.ManualActivity"
                $srObject | Add-Member -MemberType NoteProperty -Name "SequenceId" -Value $_.SequenceId
                $srObject | Add-Member -MemberType NoteProperty -Name "ChildId" -Value $_.ChildId
                $srObject | Add-Member -MemberType NoteProperty -Name "Notes" -Value $_.Notes
                $srObject | Add-Member -MemberType NoteProperty -Name "Status" -Value $_.Status
                $srObject | Add-Member -MemberType NoteProperty -Name "StatusDisplayName" -Value $_.Status.DisplayName
                $srObject | Add-Member -MemberType NoteProperty -Name "Priority" -Value $_.Priority
                $srObject | Add-Member -MemberType NoteProperty -Name "PriorityDisplayName" -Value $_.Priority.DisplayName
                $srObject | Add-Member -MemberType NoteProperty -Name "Area" -Value $_.Area
                $srObject | Add-Member -MemberType NoteProperty -Name "AreaDisplayName" -Value $_.Area.DisplayName
                $srObject | Add-Member -MemberType NoteProperty -Name "Stage" -Value $_.Stage
                $srObject | Add-Member -MemberType NoteProperty -Name "StageDisplayName" -Value $_.Stage.DisplayName
                $srObject | Add-Member -MemberType NoteProperty -Name "Documentation" -Value $_.Documentation
                $srObject | Add-Member -MemberType NoteProperty -Name "Skip" -Value ("$" + $_.Skip)
                $srObject | Add-Member -MemberType NoteProperty -Name "Id" -Value $_.Id
                $srObject | Add-Member -MemberType NoteProperty -Name "Title" -Value $_.Title
                $srObject | Add-Member -MemberType NoteProperty -Name "Description" -Value $_.Description
                $srObject | Add-Member -MemberType NoteProperty -Name "ContactMethod" -Value $_.ContactMethod
                $srObject | Add-Member -MemberType NoteProperty -Name "CreatedDate" -Value $_.CreatedDate
                $srObject | Add-Member -MemberType NoteProperty -Name "ScheduledStartDate" -Value $_.ScheduledStartDate
                $srObject | Add-Member -MemberType NoteProperty -Name "ScheduledEndDate" -Value $_.ScheduledEndDate
                $srObject | Add-Member -MemberType NoteProperty -Name "ActualStartDate" -Value $_.ActualStartDate
                $srObject | Add-Member -MemberType NoteProperty -Name "ActualEndDate" -Value $_.ActualEndDate
                $srObject | Add-Member -MemberType NoteProperty -Name "IsDowntime" -Value ("$" + $_.IsDowntime)
                $srObject | Add-Member -MemberType NoteProperty -Name "IsParent" -Value ("$" + $_.IsParent)
                $srObject | Add-Member -MemberType NoteProperty -Name "ScheduledDowntimeStartDate" -Value $_.ScheduledDowntimeStartDate
                $srObject | Add-Member -MemberType NoteProperty -Name "ScheduledDowntimeEndDate" -Value $_.ScheduledDowntimeEndDate
                $srObject | Add-Member -MemberType NoteProperty -Name "ActualDowntimeStartDate" -Value $_.ActualDowntimeStartDate
                $srObject | Add-Member -MemberType NoteProperty -Name "ActualDowntimeEndDate" -Value $_.ActualDowntimeEndDate
                $srObject | Add-Member -MemberType NoteProperty -Name "RequiredBy" -Value $_.RequiredBy
                $srObject | Add-Member -MemberType NoteProperty -Name "PlannedCost" -Value $_.PlannedCost
                $srObject | Add-Member -MemberType NoteProperty -Name "ActualCost" -Value $_.ActualCost
                $srObject | Add-Member -MemberType NoteProperty -Name "PlannedWork" -Value $_.PlannedWork
                $srObject | Add-Member -MemberType NoteProperty -Name "ActualWork" -Value $_.ActualWork
                $srObject | Add-Member -MemberType NoteProperty -Name "UserInput" -Value $_.UserInput
                $srObject | Add-Member -MemberType NoteProperty -Name "FirstAssignedDate" -Value $_.FirstAssignedDate
                $srObject | Add-Member -MemberType NoteProperty -Name "FirstResponseDate" -Value $_.FirstResponseDate
                $srObject | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $_.DisplayName
                $srObject | Add-Member -MemberType NoteProperty -Name "Name" -Value $_.Name
                $srObject | Add-Member -MemberType NoteProperty -Name "Path" -Value $_.Path
                $srObject | Add-Member -MemberType NoteProperty -Name "FullName" -Value $_.FullName
                $srObject | Add-Member -MemberType NoteProperty -Name "TimeAdded" -Value $_.TimeAdded
                $srObject | Add-Member -MemberType NoteProperty -Name "LastModifiedBy" -Value $_.LastModifiedBy
                $srObject | Add-Member -MemberType NoteProperty -Name "LastModified" -Value $_.LastModified
                $srObject | Add-Member -MemberType NoteProperty -Name "IsNew" -Value ("$" + $_.IsNew)
                $srObject | Add-Member -MemberType NoteProperty -Name "HasChanges" -Value ("$" + $_.HasChanges)
                $srObject | Add-Member -MemberType NoteProperty -Name "GroupsAsDifferentType" -Value ("$" + $_.GroupsAsDifferentType)
                $srObject | Add-Member -MemberType NoteProperty -Name "ViewName" -Value $_.ViewName
                $srObject | Add-Member -MemberType NoteProperty -Name "ObjectMode" -Value $_.ObjectMode
                $srArray += $srObject
            }
        }        
        $i++
    }

    $srArray | Export-Csv $filePath
    $srObject = $null
    $srArray = $null
}

function Export-RelatedReviewActivity([string]$filePath)
{
    $i = 0
    $class = Get-SCSMClass -Name System.WorkItem.ServiceRequest$
    $userClass = Get-SCSMClass -Name System.Domain.User$

    $srArray =@()

    $totalCount = (Get-SCSMObject -Class $class).Count

    Get-SCSMObject -Class $class | ForEach-Object {
        $id = $_.Id
        $raCount = (Get-SCSMRelatedObject -SMObject $_ | Where-Object {$_.ClassName -eq "System.WorkItem.Activity.ReviewActivity"}).Count
        Write-Verbose "$id has $raCount Review Activities"
        Write-Progress -Activity "Exporting Related Review Activities" -Status "Exported $raCount Review Activities which were related to $id" -PercentComplete (($i / $totalCount) * 100)
        Get-SCSMRelatedObject -SMObject $_ | Where-Object {$_.ClassName -eq "System.WorkItem.Activity.ReviewActivity"} | ForEach-Object {
            
            $srObject = New-Object PSObject
            $srObject | Add-Member -MemberType NoteProperty -Name "Parent" -Value $id
            $srObject | Add-Member -MemberType NoteProperty -Name "ApprovalCondition" -Value $_.ApprovalCondition
            $srObject | Add-Member -MemberType NoteProperty -Name "ApprovalConditionDisplayName" -Value $_.ApprovalCondition.DisplayName
            $srObject | Add-Member -MemberType NoteProperty -Name "ApprovalPercentage" -Value $_.ApprovalPercentage
            $srObject | Add-Member -MemberType NoteProperty -Name "Comments" -Value $_.Comments
            $srObject | Add-Member -MemberType NoteProperty -Name "LineManagerShouldReview" -Value $_.LineManagerShouldReview
            $srObject | Add-Member -MemberType NoteProperty -Name "OwnersOfConfigItemShouldReview" -Value $_.OwnersOfConfigItemShouldReview
            $srObject | Add-Member -MemberType NoteProperty -Name "SequenceId" -Value $_.SequenceId
            $srObject | Add-Member -MemberType NoteProperty -Name "ChildId" -Value $_.ChildId
            $srObject | Add-Member -MemberType NoteProperty -Name "Notes" -Value $_.Notes
            $srObject | Add-Member -MemberType NoteProperty -Name "Status" -Value $_.Status
            $srObject | Add-Member -MemberType NoteProperty -Name "StatusDisplayName" -Value $_.Status.DisplayName
            $srObject | Add-Member -MemberType NoteProperty -Name "Priority" -Value $_.Priority
            $srObject | Add-Member -MemberType NoteProperty -Name "PriorityDisplayName" -Value $_.Priority.DisplayName
            $srObject | Add-Member -MemberType NoteProperty -Name "Area" -Value $_.Area
            $srObject | Add-Member -MemberType NoteProperty -Name "AreaDisplayName" -Value $_.Area.DisplayName
            $srObject | Add-Member -MemberType NoteProperty -Name "Stage" -Value $_.Stage
            $srObject | Add-Member -MemberType NoteProperty -Name "StageDisplayName" -Value $_.Stage.DisplayName
            $srObject | Add-Member -MemberType NoteProperty -Name "Documentation" -Value $_.Documentation
            $srObject | Add-Member -MemberType NoteProperty -Name "Skip" -Value $_.Skip
            $srObject | Add-Member -MemberType NoteProperty -Name "Id" -Value $_.Id
            $srObject | Add-Member -MemberType NoteProperty -Name "Title" -Value $_.Title
            $srObject | Add-Member -MemberType NoteProperty -Name "Description" -Value $_.Description
            $srObject | Add-Member -MemberType NoteProperty -Name "ContactMethod" -Value $_.ContactMethod
            $srObject | Add-Member -MemberType NoteProperty -Name "CreatedDate" -Value $_.CreatedDate
            $srObject | Add-Member -MemberType NoteProperty -Name "ScheduledStartDate" -Value $_.ScheduledStartDate
            $srObject | Add-Member -MemberType NoteProperty -Name "ScheduledEndDate" -Value $_.ScheduledEndDate
            $srObject | Add-Member -MemberType NoteProperty -Name "ActualStartDate" -Value $_.ActualStartDate
            $srObject | Add-Member -MemberType NoteProperty -Name "ActualEndDate" -Value $_.ActualEndDate
            $srObject | Add-Member -MemberType NoteProperty -Name "IsDowntime" -Value $_.IsDowntime
            $srObject | Add-Member -MemberType NoteProperty -Name "IsParent" -Value $_.IsParent
            $srObject | Add-Member -MemberType NoteProperty -Name "ScheduledDowntimeStartDate" -Value $_.ScheduledDowntimeStartDate
            $srObject | Add-Member -MemberType NoteProperty -Name "ScheduledDowntimeEndDate" -Value $_.ScheduledDowntimeEndDate
            $srObject | Add-Member -MemberType NoteProperty -Name "ActualDowntimeStartDate" -Value $_.ActualDowntimeStartDate
            $srObject | Add-Member -MemberType NoteProperty -Name "ActualDowntimeEndDate" -Value $_.ActualDowntimeEndDate
            $srObject | Add-Member -MemberType NoteProperty -Name "RequiredBy" -Value $_.RequiredBy
            $srObject | Add-Member -MemberType NoteProperty -Name "PlannedCost" -Value $_.PlannedCost
            $srObject | Add-Member -MemberType NoteProperty -Name "ActualCost" -Value $_.ActualCost
            $srObject | Add-Member -MemberType NoteProperty -Name "PlannedWork" -Value $_.PlannedWork
            $srObject | Add-Member -MemberType NoteProperty -Name "ActualWork" -Value $_.ActualWork
            $srObject | Add-Member -MemberType NoteProperty -Name "UserInput" -Value $_.UserInput
            $srObject | Add-Member -MemberType NoteProperty -Name "FirstAssignedDate" -Value $_.FirstAssignedDate
            $srObject | Add-Member -MemberType NoteProperty -Name "FirstResponseDate" -Value $_.FirstResponseDate
            $srObject | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $_.DisplayName
            $srObject | Add-Member -MemberType NoteProperty -Name "Name" -Value $_.Name
            $srObject | Add-Member -MemberType NoteProperty -Name "Path" -Value $_.Path
            $srObject | Add-Member -MemberType NoteProperty -Name "FullName" -Value $_.FullName
            $srObject | Add-Member -MemberType NoteProperty -Name "TimeAdded" -Value $_.TimeAdded
            $srObject | Add-Member -MemberType NoteProperty -Name "LastModifiedBy" -Value $_.LastModifiedBy
            $srObject | Add-Member -MemberType NoteProperty -Name "LastModified" -Value $_.LastModified
            $srObject | Add-Member -MemberType NoteProperty -Name "IsNew" -Value $_.IsNew
            $srObject | Add-Member -MemberType NoteProperty -Name "HasChanges" -Value $_.HasChanges
            $srObject | Add-Member -MemberType NoteProperty -Name "GroupsAsDifferentType" -Value ("$" + $_.GroupsAsDifferentType)
            $srObject | Add-Member -MemberType NoteProperty -Name "ViewName" -Value $_.ViewName
            $srObject | Add-Member -MemberType NoteProperty -Name "ObjectMode" -Value $_.ObjectMode
            $srArray += $srObject
        }
        
        $i++
    }

    $srArray | Export-Csv $filePath
    $srObject = $null
    $srArray = $null
}

function Export-ActivityLog([string]$filePath)
{
    $i = 0
    $class = Get-SCSMClass -Name System.WorkItem.ServiceRequest$
    $userClass = Get-SCSMClass -Name System.Domain.User$  
    $analystComment = Get-SCSMRelationshipClass -Name System.WorkItem.TroubleTicket.AnalystCommentLog$
    $userComment  = Get-SCSMRelationshipClass -Name System.WorkItem.TroubleTicket.UserCommentLog$ 

    $srArray =@()

    $totalCount = (Get-SCSMObject -Class $class).Count

    Get-SCSMObject -Class $class | ForEach-Object {
        $id = $_.Id
        $alCount = (Get-SCSMRelatedObject -SMObject $_ | Where-Object {$_.ClassName -eq "System.WorkItem.TroubleTicket.UserCommentLog" -or $_.ClassName -eq "System.WorkItem.TroubleTicket.AnalystCommentLog"}).Count
        Write-Progress -Activity "Exporting Activity Log" -Status "Exporting $id which has $alCount Activity Log entries" -PercentComplete (($i / $totalCount) * 100)
        Write-Verbose "$id has $alCount items in the Activity Log"
        Get-SCSMRelatedObject -SMObject $_ | Where-Object {$_.ClassName -eq "System.WorkItem.TroubleTicket.UserCommentLog" -or $_.ClassName -eq "System.WorkItem.TroubleTicket.AnalystCommentLog"} | ForEach-Object {
            $isPrivate = $_.IsPrivate
            $srObject = New-Object PSObject
            $srObject | Add-Member -MemberType NoteProperty -Name "RelatedServiceRequest" -Value $id
            $srObject | Add-Member -MemberType NoteProperty -Name "EnteredDate" -Value $_.EnteredDate
            $srObject | Add-Member -MemberType NoteProperty -Name "EnteredBy" -Value $_.EnteredBy
            $srObject | Add-Member -MemberType NoteProperty -Name "Comment" -Value $_.Comment
            switch($_.ClassName) {
                "System.WorkItem.TroubleTicket.AnalystCommentLog" {
                    $srObject | Add-Member -MemberType NoteProperty -Name "IsPrivate" -Value $isPrivate
                    $srObject | Add-Member -MemberType NoteProperty -Name "LogType" -value "AnalystComment"
                }
                "System.WorkItem.TroubleTicket.UserCommentLog"{
                    $srObject | Add-Member -MemberType NoteProperty -Name "LogType" -Value "UserComment"
                }
            }

            $srArray += $srObject
        }

        $i++
        
    }

    $srArray | Export-Csv $filePath
    $srObject = $null
    $srArray = $null
}

Import-SMLets
Export-ServiceRequest $csvOutput
Export-RelatedManualActivity $csvOutputRelatedManualActivities
Export-RelatedManualActivityInsideAParallelActivity $csvOutputRelatedParallelActivities
Export-RelatedReviewActivity $csvOutputRelatedReviewActivities
Export-ActivityLog $csvOutputActivityLog