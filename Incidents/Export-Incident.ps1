<#
.SYNOPSIS
    Exports Incidents from System Center Service Manager to a CSV for simplified importing to other environments

.DESCRIPTION
    Exports all Incidents and their useful/related data to a CSV file, which can be used for good or for awesome

    The purpose of this script is to help automate the process of migrating to a new Service Manager environment, bringing
    only the data you want. In this script's current life, we grab all incidents that can be found using Get-SCSMObject and
    write out two files:
        - A CSV containing all incidents, their data, and their assigned to/affected by related users
        - A CSV containing all activity log entries for all incidents
    With these CSVs, you can re-import into a new environment, or quickly populate a development environment

.PARAMETER csvOutput
    The path and name of a CSV file that is created containing all Incident data
    
    This field is required, and should be a location that you have write access to

.PARAMETER csvOutputActivityLog
    The path and name of a CSV file that is created containing all Incident Activity Log data

    This field is required, and should be a location that you have write access to

.EXAMPLE
    Export all Incidents from a Service Manager environment

    Export-Incident -csvOutput .\incidents.csv -csvOutputActivityLog .\incidents_activitylogs.csv

.NOTES
    Requires the Windows Management Framework 4 or greater
    Requires the SMlets PowerShell module
    Tested with System Center Service Manager 2012 R2 and 2016

    Jacob Thornberry
    @altrhombus
    Made with ❤ in MKE
#>

[CmdletBinding()]
Param (
    [Parameter(Mandatory = $True, HelpMessage = "Where should we write out the list of Incidents in CSV format")]
    $csvOutput,

    [Parameter(Mandatory = $True, HelpMessage = "Where should we write out the Incidents Activity Log CSV")]
    $csvOutputActivityLog
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

function Export-Incident([string]$filePath)
{
    Write-Progress -Activity "Initializing" -Status "Reticulating Splines"
    $i = 0
    $id = 0
    $class = Get-SCSMClass -Name System.WorkItem.Incident$
    $userClass = Get-SCSMClass -Name System.Domain.User$  
    $affectedUserRelClass = Get-SCSMRelationshipClass -Name System.WorkItemAffectedUser$
    $assignedToUserRelClass  = Get-SCSMRelationshipClass -Name System.WorkItemAssignedToUser$ 

    $irArray =@()

    $totalCount = (Get-SCSMObject -Class $class).Count

    Get-SCSMObject -Class $class | ForEach-Object {
        $id = $_.Id
        $affectedUser = Get-SCSMRelatedObject -SMObject $_ -Relationship $affectedUserRelClass
        $assignedTo = Get-SCSMRelatedObject -SMObject $_ -Relationship $assignedToUserRelClass

        Write-Progress -Activity "Exporting Incidents" -Status "Found $id which affects $affectedUser and was assigned to $assignedTo" -PercentComplete (($i / $totalCount) * 100)
        
        $irObject = New-Object PSObject
        $irObject | Add-Member -MemberType NoteProperty -Name "TargetResolutionTime" -Value $_.TargetResolutionTime
        $irObject | Add-Member -MemberType NoteProperty -Name "Escalated" -Value $_.Escalated
        $irObject | Add-Member -MemberType NoteProperty -Name "Source" -Value $_.Source
        $irObject | Add-Member -MemberType NoteProperty -Name "SourceDisplayName" -Value $_.Source.DisplayName
        $irObject | Add-Member -MemberType NoteProperty -Name "Status" -Value $_.Status
        $irObject | Add-Member -MemberType NoteProperty -Name "StatusDisplayName" -Value $_.Status.DisplayName
        $irObject | Add-Member -MemberType NoteProperty -Name "ResolutionDescription" -Value $_.ResolutionDescription
        $irObject | Add-Member -MemberType NoteProperty -Name "NeedsKnowledgeArticle" -Value $_.NeedsKnowledgeArticle
        $irObject | Add-Member -MemberType NoteProperty -Name "TierQueue" -Value $_.TierQueue
        $irObject | Add-Member -MemberType NoteProperty -Name "TierQueueDisplayName" -Value $_.TierQueue.DisplayName
        $irObject | Add-Member -MemberType NoteProperty -Name "HasCreatedKnowledgeArticle" -Value $_.HasCreatedKnowledgeArticle
        $irObject | Add-Member -MemberType NoteProperty -Name "LastModifiedSource" -Value $_.LastModifiedSource
        $irObject | Add-Member -MemberType NoteProperty -Name "Classification" -Value $_.Classification
        $irObject | Add-Member -MemberType NoteProperty -Name "ClassificationDisplayName" -Value $_.Classification.DisplayName
        $irObject | Add-Member -MemberType NoteProperty -Name "ResolutionCategory" -Value $_.ResolutionCategory
        $irObject | Add-Member -MemberType NoteProperty -Name "ResolutionCategoryDisplayName" -Value $_.ResolutionCategory.DisplayName
        $irObject | Add-Member -MemberType NoteProperty -Name "Priority" -Value $_.Priority
        $irObject | Add-Member -MemberType NoteProperty -Name "Impact" -Value $_.Impact
        $irObject | Add-Member -MemberType NoteProperty -Name "Urgency" -Value $_.Urgency
        $irObject | Add-Member -MemberType NoteProperty -Name "ClosedDate" -Value $_.ClosedDate
        $irObject | Add-Member -MemberType NoteProperty -Name "ResolvedDate" -Value $_.ResolvedDate
        $irObject | Add-Member -MemberType NoteProperty -Name "Id" -Value $_.Id
        $irObject | Add-Member -MemberType NoteProperty -Name "Title" -Value $_.Title
        $irObject | Add-Member -MemberType NoteProperty -Name "Description" -Value $_.Description
        $irObject | Add-Member -MemberType NoteProperty -Name "ContactMethod" -Value $_.ContactMethod
        $irObject | Add-Member -MemberType NoteProperty -Name "CreatedDate" -Value $_.CreatedDate
        $irObject | Add-Member -MemberType NoteProperty -Name "ScheduledStartDate" -Value $_.ScheduledStartDate
        $irObject | Add-Member -MemberType NoteProperty -Name "ScheduledEndDate" -Value $_.ScheduledEndDate
        $irObject | Add-Member -MemberType NoteProperty -Name "ActualStartDate" -Value $_.ActualStartDate
        $irObject | Add-Member -MemberType NoteProperty -Name "ActualEndDate" -Value $_.ActualEndDate
        $irObject | Add-Member -MemberType NoteProperty -Name "IsDowntime" -Value $_.IsDowntime
        $irObject | Add-Member -MemberType NoteProperty -Name "IsParent" -Value $_.IsParent
        $irObject | Add-Member -MemberType NoteProperty -Name "ScheduledDowntimeStartDate" -Value $_.ScheduledDowntimeStartDate
        $irObject | Add-Member -MemberType NoteProperty -Name "ScheduledDowntimeEndDate" -Value $_.ScheduledDowntimeEndDate
        $irObject | Add-Member -MemberType NoteProperty -Name "ActualDowntimeStartDate" -Value $_.ActualDowntimeStartDate
        $irObject | Add-Member -MemberType NoteProperty -Name "ActualDowntimeEndDate" -Value $_.ActualDowntimeEndDate
        $irObject | Add-Member -MemberType NoteProperty -Name "RequiredBy" -Value $_.RequiredBy
        $irObject | Add-Member -MemberType NoteProperty -Name "PlannedCost" -Value $_.PlannedCost
        $irObject | Add-Member -MemberType NoteProperty -Name "ActualCost" -Value $_.ActualCost
        $irObject | Add-Member -MemberType NoteProperty -Name "PlannedWork" -Value $_.PlannedWork
        $irObject | Add-Member -MemberType NoteProperty -Name "ActualWork" -Value $_.ActualWork
        $irObject | Add-Member -MemberType NoteProperty -Name "UserInput" -Value $_.UserInput
        $irObject | Add-Member -MemberType NoteProperty -Name "FirstAssignedDate" -Value $_.FirstAssignedDate
        $irObject | Add-Member -MemberType NoteProperty -Name "FirstResponseDate" -Value $_.FirstResponseDate
        $irObject | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $_.DisplayName
        $irObject | Add-Member -MemberType NoteProperty -Name "Name" -Value $_.Name
        $irObject | Add-Member -MemberType NoteProperty -Name "TimeAdded" -Value $_.TimeAdded
        $irObject | Add-Member -MemberType NoteProperty -Name "AffectedUser" -Value $affectedUser
        $irObject | Add-Member -MemberType NoteProperty -Name "AssignedTo" -Value $assignedTo

        $irArray += $irObject
        $i++
    }

    $irArray | Export-Csv $filePath
    
    $irObject = $null
    $irArray = $null
}

function Export-ActivityLog([string]$filePath)
{
    $i = 0
    $class = Get-SCSMClass -Name System.WorkItem.Incident$
    $userClass = Get-SCSMClass -Name System.Domain.User$  
    $analystComment = Get-SCSMRelationshipClass -Name System.WorkItem.TroubleTicket.AnalystCommentLog$
    $userComment  = Get-SCSMRelationshipClass -Name System.WorkItem.TroubleTicket.UserCommentLog$ 

    $irArray =@()

    $totalCount = (Get-SCSMObject -Class $class).Count

    Get-SCSMObject -Class $class | ForEach-Object {
        $id = $_.Id
        $alCount = (Get-SCSMRelatedObject -SMObject $_ | Where-Object {$_.ClassName -eq "System.WorkItem.TroubleTicket.UserCommentLog" -or $_.ClassName -eq "System.WorkItem.TroubleTicket.AnalystCommentLog"}).Count
        Write-Progress -Activity "Exporting Activity Log" -Status "Exporting $id which has $alCount Activity Log entries" -PercentComplete (($i / $totalCount) * 100)
        Write-Verbose "$id has $alCount items in the Activity Log"
        Get-SCSMRelatedObject -SMObject $_ | Where-Object {$_.ClassName -eq "System.WorkItem.TroubleTicket.UserCommentLog" -or $_.ClassName -eq "System.WorkItem.TroubleTicket.AnalystCommentLog"} | ForEach-Object {
            $isPrivate = $_.IsPrivate
            $irObject = New-Object PSObject
            $irObject | Add-Member -MemberType NoteProperty -Name "RelatedIncident" -Value $id
            $irObject | Add-Member -MemberType NoteProperty -Name "EnteredDate" -Value $_.EnteredDate
            $irObject | Add-Member -MemberType NoteProperty -Name "EnteredBy" -Value $_.EnteredBy
            $irObject | Add-Member -MemberType NoteProperty -Name "Comment" -Value $_.Comment
            switch($_.ClassName) {
                "System.WorkItem.TroubleTicket.AnalystCommentLog" {
                    $irObject | Add-Member -MemberType NoteProperty -Name "IsPrivate" -Value $isPrivate
                    $irObject | Add-Member -MemberType NoteProperty -Name "LogType" -value "AnalystComment"
                }
                "System.WorkItem.TroubleTicket.UserCommentLog"{
                    $irObject | Add-Member -MemberType NoteProperty -Name "LogType" -Value "UserComment"
                }
            }

            $irArray += $irObject
        }

        $i++
        
    }

    $irArray | Export-Csv $filePath
    $irObject = $null
    $irArray = $null
}

Import-SMLets
Export-Incident $CSVOutput
Export-ActivityLog $CSVOutputActivityLog