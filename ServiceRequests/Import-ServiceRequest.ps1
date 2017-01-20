
<#
.SYNOPSIS
    Imports Service Requests and related Activities to System Center Service Manager from a CSV

.DESCRIPTION
    Imports Service Requests from a CSV file (probably created by the included Export-ServiceRequest.ps1 script)
    During the Service Request import process, we also import related Manual, Review, and Parallel Activities.

    The purpose of this script is to help automate the process of migrating to a new Service Manager environment, bringing
    only the data you want. In this script's current life, we import all Service Requests that are contained in your CSV.
    We require the following:
        - A CSV containing all the service requests and their supporting data
        - A CSV containing the service requests' activity log data
        - A CSV containing the service requests' related Manual Activities
        - A CSV containing the service requests' related Review Activities
        - A CSV containing the service requests' related Parallel Activities
        - A "surrogate" user in the event that an assigned/affected user no longer exists
        - A directory containing all service request attachments
        - Write access to a directory to create a temporary "diffTable" CSV. By default, this script recreates all service
        requests using the same SR numbers as before. If this is a lab (or a production environment that you purposefully want
        to start with new numbers), this CSV is used to keep track of how the SR number is related to it's previous (old) 
        SR number
    
    Before importing data, we run a "sanity check" against list enums for Service Requests. If your Service Request lists don't
    contain everything your previous environment had, we'll terminate early. Therefore, we recommend that you import any required
    Management Packs prior to running this script

.PARAMETER sourceServiceRequestCsv
    The path and name of a CSV file that contains all Service Request data
    
    This field is required, and should contain the path to the CSV

.PARAMETER surrogateAffectedUser
    A user (in "Display Name" format) that will be assigned the "Affected User" of any service requests that contain an affected 
    user who no longer exists

    This field is required. If you don't tell Service Manager to sync with AD first, all of your Service Requests will likely have
    their affected user set to this person

.PARAMETER surrogateAssignedTo
    A user (in "Display Name" format) that will be assigned the "Assigned To" of any service requests that contain an AssignedTo
    user who no longer exists

    This field is required. If you don't tell Service Manager to sync with AD first, all of your Service Requests will likely be 
    assigned to this person. If this is a big import, you should get them a coffee

.PARAMETER sourceFileAttachmentsDir
    The path of a directory that contains Service Request attachments

    This field is not required, and should be a location that you have read access to. This directory should contain a list of directories
    named by ID (SR1, etc.). We expect to find the SR's attachments in these subdirectories

.PARAMETER sourceActivityLogCsv
    The path and name of a CSV file that contains all Service Request Activity Log data
    
    This field is required, and should contain the path to the CSV

.PARAMETER sourceManualActivityCsv
    The path and name of a CSV file that contains all related Manual Activities for your Service Requests

    This field is required, and should contain the path to the CSV

.PARAMETER sourceReviewActivityCsv
    The path and name of a CSV file that contains all related Review Activities for your Service Requests

    This field is required, and should contain the path to the CSV

.PARAMETER sourceParallelActivityCsv
    The path and name of a CSV file that contains all related Parallel Activities (and their related Manual and Review Activities) for 
    your Service Requests

    The code that processes these is my least favorite of all, and is on deck to be swiftly rewritten. Until then, please note that we
    only go "one level deep" for parallel activities. If you have a PA with RA's and MA's under it, we'll import them fine. If you have
    a PA and another PA or SA underneath it, you're going to have a bad time.

    This field is required, and should contain the path to the CSV

.PARAMETER diffTableCsv
    The path and name of a CSV file that the script will write out to keep track of existing and new service request ID information
    
    This field is required, and should contain the path to the CSV
    This path needs to be writable by the account running this script

.EXAMPLE
    Export all Service Requests from a Service Manager environment

    .\Import-ServiceRequest.ps1 -sourceServiceRequestCsv .\sr_output.csv -surrogateAffectedUser "Douglas, Fred" -surrogateAssignedTo "Douglas, Fred" -diffTableCsv .\difftable.csv -sourceFileAttachmentsDir .\exported_SR_attachments -sourceActivityLogCsv .\sr_al_output.csv -sourceManualActivityCsv .\sr_ma_output.csv -sourceReviewActivityCsv .\sr_ra_output.csv -sourceParallelActivityCsv .\sr_pa_output.csv

.NOTES
    Requires the Windows Management Framework 4 or greater
    Requires the SMlets PowerShell module
    Tested with System Center Service Manager 2012 R2 and 2016

    Please be advised of the following known issues/limitations:
        - We only import Parallel Activities and their related activities one level deep.
        - Imported SRs without an ImplementationResults are being set to "Partially Implemented"
        - The Activity Implementer property is not being assigned for Manual Activities
        - No reviewer/voter information is assigned for Review Activities

    Jacob Thornberry
    @jakertberry
#>

[CmdletBinding()]
Param (
    [Parameter(Mandatory = $True, HelpMessage = "The CSV containing the service requests you want to import")]
    [string]$sourceServiceRequestCsv,

    [Parameter(Mandatory = $True, HelpMessage = "The user in DisplayName format that will be used if the affected user no longer exists")]
    [string]$surrogateAffectedUser,

    [Parameter(Mandatory = $True, HelpMessage = "The user in DisplayName format that will be used if the assigned to user no longer exists")]
    [string]$surrogateAssignedTo,

    [Parameter(Mandatory = $True, HelpMessage = "Directory containing previously-exported attachments. Will skip attachments if not specified")]
    [string]$sourceFileAttachmentsDir,

    [Parameter(Mandatory = $True, HelpMessage = "CSV containing previously-exported activity log entries. Will skip if not specified")]
    [string]$sourceActivityLogCsv,

    [Parameter(Mandatory = $True, HelpMessage = "CSV containing previously-exported related Manual Activities. Will skip if not specified")]
    [string]$sourceManualActivityCsv,

    [Parameter(Mandatory = $True, HelpMessage = "CSV containing previously-exported related Review Activities entries. Will skip if not specified")]
    [string]$sourceReviewActivityCsv,

    [Parameter(Mandatory = $True, HelpMessage = "CSV containing previously-exported related Parallel Activities entries. Will skip if not specified")]
    [string]$sourceParallelActivityCsv,

    [Parameter(Mandatory = $True, HelpMessage = "This script needs to write out a file that lists your old SR numbers alongside your new ones. Enter the full path here.")]
    [string]$diffTableCsv
)

$srDiffTableArray=@()

function Import-SMLets
{
    Write-Progress -Activity "Initializing" -Status "Checking for SMLets"
    if (!(Get-Module SMLets))
    {
        Write-Progress -Activity "Initializing" -Status "Importing the SMLets Module"
        Import-Module SMLets -Force -ErrorAction Stop
    }
}

function Get-Down
{
    $magic = 5
    for ($i=$magic; $i -gt 0; $i--) {
        Write-Progress -Activity "Take a deep breath and ready your wand--its almost time for some magic!" -Status "Reticulating splines" -SecondsRemaining $i
        Start-Sleep 1
    }
}

function Test-EnumSanity([string]$filePath)
{
    $serviceRequestClass = Get-SCSMClass -Name System.WorkItem.ServiceRequest$
    $userClass = Get-SCSMClass -Name System.Domain.User$  
    $serviceRequestAffectedUserRelClass = Get-SCSMRelationshipClass -Name System.WorkItemAffectedUser$
    $assignedToUserRelClass  = Get-SCSMRelationshipClass -Name System.WorkItemAssignedToUser$ 
    $stopCondition = $false
    $i = 0
    $totalCount = (Import-Csv $filePath).Count

    Import-Csv $filePath | ForEach-Object {
        Write-Progress -Activity "Verifying environment Enum values" -Status "Checking existing service request $($_.Id)" -PercentComplete (($i / $totalCount) * 100)
        
        $testStatusDisplayName = $_.StatusDisplayName
        $testPriorityDisplayName = $_.PriorityDisplayName
        $testUrgencyDisplayName = $_.UrgencyDisplayName
        $testSourceDisplayName = $_.SourceDisplayName
        $testImplementationResultsDisplayName = $_.ImplementationResultsDisplayName
        $testAreaDisplayName = $_.AreaDisplayName
        $testSupportGroupDisplayName = $_.SupportGroupDisplayName

        if (((Get-SCSMEnumeration | Where-Object { $_.DisplayName -eq $testStatusDisplayName }).Count -eq 0) -xor ($testStatusDisplayName -eq ""))
        {
            Write-Error -Message "The enum value $testStatusDisplayName for Status does not exist in this environment." -ErrorAction Continue
            $stopCondition = $true
        }
        if (((Get-SCSMEnumeration | Where-Object { $_.DisplayName -eq $testPriorityDisplayName }).Count -eq 0) -xor ($testPriorityDisplayName -eq ""))
        {
            Write-Error -Message "The enum value $testPriorityDisplayName for Priority does not exist in this environment." -ErrorAction Continue
            $stopCondition = $true
        }
        if (((Get-SCSMEnumeration | Where-Object { $_.DisplayName -eq $testUrgencyDisplayName }).Count -eq 0) -xor ($testUrgencyDisplayName -eq ""))
        {
            Write-Error -Message "The enum value $testUrgencyDisplayName for Urgency does not exist in this environment." -ErrorAction Continue
            $stopCondition = $true
        }
        if (((Get-SCSMEnumeration | Where-Object { $_.DisplayName -eq $testSourceDisplayName }).Count -eq 0) -xor ($testSourceDisplayName -eq ""))
        {
            Write-Error -Message "The enum value $testSourceDisplayName for Source does not exist in this environment." -ErrorAction Continue
            $stopCondition = $true
        }
        if (((Get-SCSMEnumeration | Where-Object { $_.DisplayName -eq $testImplementationResultsDisplayName }).Count -eq 0) -xor ($testImplementationResultsDisplayName -eq ""))
        {
            Write-Error -Message "The enum value $testImplementationResultsDisplayName for ImplementationResults does not exist in this environment." -ErrorAction Continue
            $stopCondition = $true
        }
        if (((Get-SCSMEnumeration | Where-Object { $_.DisplayName -eq $testAreaDisplayName }).Count -eq 0) -xor ($testAreaDisplayName -eq ""))
        {
            Write-Error -Message "The enum value $testAreaDisplayName for DisplayName does not exist in this environment." -ErrorAction Continue
            $stopCondition = $true
        }
        if (((Get-SCSMEnumeration | Where-Object { $_.DisplayName -eq $testSupportGroupDisplayName }).Count -eq 0) -xor ($testSupportGroupDisplayName -eq ""))
        {
            Write-Error -Message "The enum value $testSupportGroupDisplayName for SupportGroup does not exist in this environment." -ErrorAction Continue
            $stopCondition = $true
        }
        $i++
    }

    if ($stopCondition -eq $true)
    {
        Write-Error "Enum value dependency failure. Please add missing values to lists or import missing Management Packs." -ErrorAction Stop
    }
}

function Import-ServiceRequest([string]$filePath, [string]$diffPath, [string]$raPath, [string]$maPath, [string]$paPath)
{
    $serviceRequestClass = Get-SCSMClass -Name System.WorkItem.ServiceRequest$
    $userClass = Get-SCSMClass -Name System.Domain.User$  
    $serviceRequestAffectedUserRelClass = Get-SCSMRelationshipClass -Name System.WorkItemAffectedUser$
    $assignedToUserRelClass  = Get-SCSMRelationshipClass -Name System.WorkItemAssignedToUser$ 
    $i = 0
    $totalCount = (Import-Csv $filePath).Count

    Import-Csv $filePath | ForEach-Object {
        $srObject = New-Object PSObject
        $srObject | Add-Member -MemberType NoteProperty -Name "PreviousId" -Value $_.Id
        $affectedUser = ""
        $assignedTo = ""

        $srId = $_.Id
        $srTitle = $_.Title
        $srDescription = $_.Description
        $srStatusDisplayName = $_.StatusDisplayName
        $srTemplateId = $_.TemplateId
        $srPriority = $_.PriorityDisplayName
        $srUrgency = $_.UrgencyDisplayName
        $srSourceDisplayName = $_.SourceDisplayName
        $srImplementationResults = $_.ImplementationResultsDisplayName
        $srNotes = $_.Notes
        $srArea = $_.AreaDisplayName
        $srSupportGroup = $_.SupportGroupDisplayName
        $srContactMethod = $_.ContactMethod
        $srUserInput = $_.UserInput
        $srViewName = $_.ViewName
        $srObjectMode = $_.ObjectMode

        #$srPlannedCost = $_.PlannedCost
        #$srActualCost = $_.ActualCost
        #$srPlannedWork = $_.PlannedWork
        #$srActualWork = $_.ActualWork

        
        if ($srUserInput.Length -eq 0) {
            $srUserInput = $null
        }
        #Change the Id below to "SR{0}" to generate a new SR instead of reuing the same SR number
        $serviceRequestHashTable = @{
            Id = $srId;
            Title = $srTitle;
            Description = $srDescription;
            Status = $srStatusDisplayName;
            TemplateId = $srTemplateId;
            Priority = $srPriority;
            Urgency = $srUrgency;
            Source = $srSourceDisplayName;
            ImplementationResults = $srImplementationResults;
            Notes = $srNotes;
            Area = $srArea;
            SupportGroup = $srSupportGroup;
            ContactMethod = $srContactMethod;
            UserInput = $srUserInput;
        }
        try {
            [datetime]$srCompletedDate = $_.CompletedDate
            $serviceRequestHashTable += @{CompletedDate = $srCompletedDate}
        } catch { }
        try {
            [datetime]$srClosedDate = $_.ClosedDate
            $serviceRequestHashTable +=@{ClosedDate = $srClosedDate}
        }
        catch { }
        try {
            [datetime]$srCreatedDate = $_.CreatedDate
            $serviceRequestHashTable +=@{CreatedDate = $srCreatedDate}
        }
        catch { }
        try {
            [datetime]$srScheduledStartDate = $_.ScheduledStartDate
            $serviceRequestHashTable +=@{ScheduledStartDate = $srScheduledStartDate}
        }
        catch { }
        try {
            [datetime]$srScheduledEndDate = $_.ScheduledEndDate
            $serviceRequestHashTable +=@{ScheduledEndDate = $srScheduledEndDate}
        }
        catch { }
        try {
            [datetime]$srActualStartDate = $_.ActualStartDate
            $serviceRequestHashTable +=@{ActualStartDate = $srActualStartDate}
        }
        catch { }
        try {
            [datetime]$srActualEndDate = $_.ActualEndDate
            $serviceRequestHashTable +=@{ActualEndDate = $srActualEndDate}
        }
        catch { }
        try {
            $srIsDowntime = $_.IsDowntime
            if($srIsDowntime -eq "True") {$srIsDowntime = $true} else {$srIsDowntime = $false}
            $serviceRequestHashTable +=@{IsDowntime = [bool]$srIsDowntime}
        }
        catch { }
        try {
            $srIsParent = $_.IsParent
            if($srIsParent -eq "True") {$srIsParent = $true} else {$srIsParent = $false}
            $serviceRequestHashTable +=@{IsParent = [bool]$srIsParent}
        }
        catch { }
        try {
            [datetime]$srScheduledDowntimeStartDate = $_.ScheduledDowntimeStartDate
            $serviceRequestHashTable +=@{ScheduledDowntimeStartDate = $srScheduledDowntimeStartDate}
        }
        catch { }
        try {
            [datetime]$srScheduledDowntimeEndDate = $_.ScheduledDowntimeEndDate
            $serviceRequestHashTable +=@{ScheduledDowntimeEndDate = $srScheduledDowntimeEndDate}
        }
        catch { }
        try {
            [datetime]$srActualDowntimeStartDate = $_.ActualDowntimeStartDate
            $serviceRequestHashTable +=@{ActualDowntimeStartDate = $srActualDowntimeStartDate}
        }
        catch { }
        try {
            [datetime]$srActualDowntimeEndDate = $_.ActualDowntimeEndDate
            $serviceRequestHashTable +=@{ActualDowntimeEndDate = $srActualDowntimeEndDate}
        }
        catch { }
        try {
            [datetime]$srScheduledDowntimeEndDate = $_.ScheduledDowntimeEndDate
            $serviceRequestHashTable +=@{ScheduledDowntimeEndDate = $srScheduledDowntimeEndDate}
        }
        catch { }
        try {
            [datetime]$srActualDowntimeStartDate = $_.ActualDowntimeStartDate
            $serviceRequestHashTable +=@{ActualDowntimeStartDate = $srActualDowntimeStartDate}
        }
        catch { }
        try {
            [datetime]$srActualDowntimeEndDate = $_.ActualDowntimeEndDate
            $serviceRequestHashTable +=@{ActualDowntimeEndDate = $srActualDowntimeEndDate}
        }
        catch { }
        try {
            [datetime]$srRequiredBy = $_.RequiredBy
            $serviceRequestHashTable +=@{RequiredBy = $srRequiredBy;}
        }
        catch { }
        try {
            [datetime]$srFirstAssignedDate = $_.FirstAssignedDate
            $serviceRequestHashTable +=@{FirstAssignedDate = $srFirstAssignedDate}
        }
        catch { }
        try {
            [datetime]$srFirstResponseDate = $_.FirstResponseDate
            $serviceRequestHashTable +=@{FirstResponseDate = $srFirstResponseDate}
        }
        catch { }
        try {
            [datetime]$srLastModified = $_.srLastModified
            $serviceRequestHashTable +=@{LastModified = $srLastModified}
        } catch { }
        try {
            [bool]$srIsNew = $_.IsNew
            if($srIsNew -eq "True") {$srIsNew = $true} else {$srIsNew = $false}
            $serviceRequestHashTable +=@{IsNew = [bool]$srIsNew}
        }
        catch { }
        try {
            [bool]$srHasChanges = $_.HasChanges
            if($srHasChanges -eq "True") {$srHasChanges = $true} else {$srHasChanges = $false}
            $serviceRequestHashTable +=@{HasChanges = [bool]$srHasChanges}
        }
        catch { }
        try {
            [bool]$srGroupsAsDifferentType = $_.GroupsAsDifferentType
            if($srGroupsAsDifferentType -eq "True") {$srGroupsAsDifferentType = $true} else {$srGroupsAsDifferentType = $false}
            $serviceRequestHashTable +=@{GroupsAsDifferentType = [bool]$srGroupsAsDifferentType}
        }
        catch { } 

        #try {
        #    [datetime]$srTimeAdded = $_.TimeAdded
        #    $serviceRequestHashTable +=@{TimeAdded = $srTimeAdded}
        #}
        #catch { }

        $surrogateAffectedUser = "insert a name here for now"
        $surrogateAssignedTo = "insert a name here for now"

        try {
            $affectedUser = Get-SCSMObject -Class $UserClass -Filter "DisplayName -eq $($_.AffectedUser)"
        } catch {
            $affectedUser = Get-SCSMObject -Class $UserClass -Filter "DisplayName -eq $($surrogateAffectedUser)"
            Write-Information "Used a surrogate Affected User for $newServiceRequestId"
        }

        try {
            $assignedTo = Get-SCSMObject -Class $UserClass -Filter "DisplayName -eq $($_.AssignedTo)"
        } catch {
            $assignedTo = Get-SCSMObject -Class $UserClass -Filter "DisplayName -eq $($surrogateAssignedTo)"
            Write-Information "Used a surrogate AssignedTo User for $newServiceRequestId"
        }

        $newServiceRequest = New-SCSMObject -Class $serviceRequestClass -PropertyHashtable $serviceRequestHashTable -PassThru
        $newServiceRequestId = $newServiceRequest.Id
        $newServiceRequestGuid = $newServiceRequest.__internalID

        $srObject | Add-Member -MemberType NoteProperty -Name "CurrentId" -Value $newServiceRequestId
        $srObject | Add-Member -MemberType NoteProperty -Name "CurrentGuid" -Value $newServiceRequestGuid
        $srDiffTableArray += $srObject

        if (($affectedUser).Count -eq 1) { New-SCSMRelationshipObject -Relationship $serviceRequestAffectedUserRelClass -Source $newServiceRequest -Target $affectedUser -Bulk }
        if (($affectedUser).Count -gt 1) { Write-Information "Unable to add the Affected User to $newServiceRequestId because there was more than one for some reason." }
        if (($assignedTo).Count -eq 1) { New-SCSMRelationshipObject -Relationship $assignedToUserRelClass -Source $newServiceRequest -Target $assignedTo -Bulk }
        if (($assignedTo).Count -gt 1) { Write-Information "Unable to add the AssignedTo User to $newServiceRequestId because there was more than one for some reason." }

        $relatedManualActivity = Import-Csv $maPath | Where-Object {$_.Parent -eq $srId}
        Import-RelatedActivity ManualActivity $relatedManualActivity $newServiceRequestGuid

        $relatedReviewActivity = Import-Csv $raPath | Where-Object {$_.Parent -eq $srId}
        Import-RelatedActivity ReviewActivity $relatedReviewActivity $newServiceRequestGuid

        $relatedParallelActivity = Import-Csv $paPath | Where-Object {$_.Parent -eq $srId}
        $relatedParallelActivityId = $relatedParallelActivity.Id

        if (!($relatedParallelActivityId -eq $null)) {
            $newRelatedParallelActivity = Import-RelatedActivity ParallelActivity $relatedParallelActivity $newServiceRequestGuid
            $relatedManualActivityInception = Import-Csv $paPath | Where-Object {$_.Parent -eq $relatedParallelActivityId}
            Import-RelatedActivity ManualActivity $relatedManualActivityInception (Get-ParallelActivityGuid $newRelatedParallelActivity)
        }


        Write-Progress -Activity "Importing Service Requests" -Status "Re-created new Service Request $newServiceRequestId from previous SR $($_.Id)" -PercentComplete (($i / $totalCount) * 100)
        $i++;
    }

    $srDiffTableArray | Export-Csv $diffPath
}

function Import-RelatedActivity([string]$activityType, $filePath, [guid]$parentSR)
{
    $i = 0
    $totalCount = ($filePath).Count

    $srClass = Get-SCSMClass -Name System.WorkItem.ServiceRequest$

    $filePath | ForEach-Object {
        switch($activityType) {
            "ReviewActivity" {
                $activityClass = Get-SCSMClass -Name System.WorkItem.Activity.ReviewActivity
                $prefixName = (Get-SCClassInstance -Class (Get-SCClass -Name System.GlobalSetting.ActivitySettings)).SystemWorkItemActivityReviewActivityIdPrefix
            }
            "ManualActivity" {
                $activityClass = Get-SCSMClass -Name System.WorkItem.Activity.ManualActivity
                $prefixName = (Get-SCClassInstance -Class (Get-SCClass -Name System.GlobalSetting.ActivitySettings)).SystemWorkItemActivityManualActivityIdPrefix
            }
            "ParallelActivity" {
                $activityClass = Get-SCSMClass -Name System.WorkItem.Activity.ParallelActivity
                $prefixName = (Get-SCClassInstance -Class (Get-SCClass -Name System.GlobalSetting.ActivitySettings)).SystemWorkItemActivityParallelActivityIdPrefix
            }
        }
        $activityRelationship = Get-SCRelationship -Name System.WorkItemContainsActivity

        # This isn't doing error checking yet for enums
        #puppymonkeybaby Id = "$prefixName{0}"
        $relatedActivityHashTable = @{
            Id = $($_.Id)
            SequenceId = $($_.SequenceId)
            Notes = $($_.Notes)
            Status = $($_.Status)
            Documentation = $($_.Documentation)
            Title = $($_.Title)
            Description = $($_.Description)
            ContactMethod = $($_.ContactMethod)
        }

        if (!($_.PriorityDisplayName.Length -eq 0)) {
            $relatedActivityHashTable +=@{Priority = $($_.PriorityDisplayName)}
        }

        if (!($_.AreaDisplayName.Length -eq 0)) {
            $relatedActivityHashTable +=@{Area = $($_.AreaDisplayName)}
        }

        if (!($_.StageDisplayName.Length -eq 0)) {
            $relatedActivityHashTable +=@{Stage = $($_.StageDisplayName)}
        }

        if (!($_.UserInput.Length -eq 0)) {
            $relatedActivityHashTable +=@{UserInput = $($_.UserInput)}
        }

        try {
            if($_.Skip -eq "True") {$skipVal = $true} else {$skipVal = $false}
            $relatedActivityHashTable +=@{Skip = $([bool]$skipVal)}
        }
        catch { }
        try {
            $relatedActivityHashTable +=@{CreatedDate = $([datetime]$_.CreatedDate)}
        }
        catch { }
        try {
            $relatedActivityHashTable +=@{ScheduledStartDate = $([datetime]$_.ScheduledDowntimeStartDate)}
        }
        catch { }
        try {
            $relatedActivityHashTable +=@{ScheduledDowntimeEndDate = $([datetime]$_.ScheduledDowntimeEndDate)}
        }
        catch { }
        try {
            $relatedActivityHashTable +=@{ActualStartDate = $([datetime]$_.ActualStartDate)}
        }
        catch { }
        try {
            $relatedActivityHashTable +=@{ActualEndDate = $([datetime]$_.ActualEndDate)}
        }
        catch { }
        try {
            if($_.IsDowntime -eq "True") {$isDowntimeVal = $true} else {$isDowntimeVal = $false}
            $relatedActivityHashTable +=@{IsDowntime = $([bool]$isDowntimeVal)}
        }
        catch { }
        try {
            if($_.IsParent -eq "True") {$isParentVal = $true} else {$isParentVal = $false}
            $relatedActivityHashTable +=@{IsParent = $([bool]$isParentVal)}
        }
        catch { }
        try {
            $relatedActivityHashTable +=@{ScheduledDowntimeStartDate = $([datetime]$_.ScheduledDowntimeStartDate)}
        }
        catch { }
        try {
            $relatedActivityHashTable +=@{ScheduledDowntimeEndDate = $([datetime]$_.ScheduledDowntimeEndDate)}
        }
        catch { }
        try {
            $relatedActivityHashTable +=@{ActualDowntimeStartDate = $([datetime]$_.ActualDowntimeStartDate)}
        }
        catch { }
        try {
            $relatedActivityHashTable +=@{ActualDowntimeEndDate = $([datetime]$_.ActualDowntimeEndDate)}
        }
        catch { }
        try {
            $relatedActivityHashTable +=@{FirstAssignedDate = $([datetime]$_.FirstAssignedDate)}
        }
        catch { }
        try {
            $relatedActivityHashTable +=@{FirstResponseDate = $([datetime]$_.FirstResponseDate)}
        }
        catch { }
        try {
            $relatedActivityHashTable +=@{ActualDowntimeStartDate = $([datetime]$_.ActualDowntimeStartDate)}
        }
        catch { }
        try {
            $relatedActivityHashTable +=@{ActualDowntimeStartDate = $([datetime]$_.ActualDowntimeStartDate)}
        }
        catch { }


        $newActivity = New-SCRelationshipInstance -RelationshipClass $activityRelationship -Source (Get-SCSMObject $parentSR) -TargetClass $activityClass -TargetProperty $relatedActivityHashTable -PassThru
        $newActivityId = $newActivity.TargetObject

        Write-Verbose "Created new activity $newActivityId with status $($_.StatusDisplayName)"
        #Write-Progress -Activity "Importing Service Request $activityType" -Status "Migrated activity log for $newActivityId" -PercentComplete (($i / $totalCount) * 100)

        if ($activityType -eq "ParallelActivity")
        {
            return $newActivityId
        }

        $i++;
        
    }

}


function Import-ServiceRequestActivityLog([string]$filePath, [string]$diffPath)
{
    $i = 0
    $totalCount = (Import-Csv $filePath).Count
    $srClass = Get-SCSMClass -Name System.WorkItem.ServiceRequest$

    Import-Csv $filePath | ForEach-Object {
        $newGUID = ([Guid]::NewGuid()).ToString()
        $oldSR = $_.RelatedServiceRequest

        $translatedFolderName = Get-ConvertedServiceRequestNumber $oldSR $diffPath
        $smObject = Get-SCSMObject -Class $srClass -Filter "DisplayName -like $translatedFolderName"

        $newComment = $_.Comment
        $newEnteredBy = $_.EnteredBy
        $newEnteredDate = $_.EnteredDate
        switch($_.LogType) {
            "AnalystComment" {
                if($_.isPrivate -eq "True") {[bool]$isPrivate = $true} else {[bool]$isPrivate = $false}
                $commentProjection = @{__CLASS = "System.WorkItem.ServiceRequest";
                                       __SEED = $smObject;
                AnalystCommentLog = @{__CLASS = "System.WorkItem.TroubleTicket.AnalystCommentLog";
                                      __OBJECT = @{Id = $newGUID;
                                                DisplayName = $newGUID;
                                                Comment = $newComment;
                                                EnteredBy  = $newEnteredBy;
                                                EnteredDate = $newEnteredDate;
                                                IsPrivate = $isPrivate
                                                }
                                    }
                }
            }
            "UserComment"{
                $commentProjection = @{__CLASS = "System.WorkItem.ServiceRequest";
                                       __SEED = $smObject;
                EndUserCommentLog = @{__CLASS = "System.WorkItem.TroubleTicket.UserCommentLog";
                                      __OBJECT = @{Id = $newGUID;
                                                DisplayName = $newGUID;
                                                Comment = $newComment;
                                                EnteredBy  = $newEnteredBy;
                                                EnteredDate = $newEnteredDate
                                            }
                                }
                }
            }
        }
        
        New-SCSMObjectProjection -Type "System.WorkItem.ServiceRequestProjection" -Projection $commentProjection
        Write-Progress -Activity "Importing Service Request Activity Log" -Status "Migrated activity log for $translatedFolderName" -PercentComplete (($i / $totalCount) * 100)
        $i++;
    }
}

function Import-ServiceRequestFileAttachments([string]$filePath, [string]$diffPath)
{
    $attachmentList = Get-ChildItem -Directory $filePath
    $i = 0
    $totalCount = $attachmentList.Count

    $managementGroup = New-Object Microsoft.EnterpriseManagement.EnterpriseManagementGroup "localhost"

    $classAL = Get-SCSMClass -Name System.WorkItem.TroubleTicket.ActionLog$
    $classFA = Get-SCSMClass -Name System.FileAttachment$

    $actionLogRel = Get-SCSMRelationshipClass -Name System.WorkItemHasActionLog$
    $fileAttachmentRel = Get-SCSMRelationshipClass -Name System.WorkItemHasFileAttachment$

    Get-ChildItem -Directory $filePath | ForEach-Object {
        $translatedFolderName = Get-ConvertedServiceRequestNumber $_.Name $diffPath
        
        Get-ChildItem -File ($_.FullName) | ForEach-Object {
            Write-Verbose "Adding $($_.FullName) to $translatedFolderName"

            $fileMode = [IO.FileMode]::Open
            $fileStream = New-Object System.IO.FileStream $($_.FullName), $fileMode

            $fileAttachment = New-Object Microsoft.EnterpriseManagement.Common.CreatableEnterpriseManagementObject($managementGroup, $classFA)

            $attachmentGuid = [Guid]::NewGuid().ToString()
            $fileAttachment.Item($classFA, "Id").Value = $attachmentGuid
            $fileAttachment.Item($classFA, "DisplayName").Value = $($_.Name)
            $fileAttachment.Item($classFA, "Description").Value = $($_.Name)
            $fileAttachment.Item($classFA, "Extension").Value = $($_.Extension)
            $fileAttachment.Item($classFA, "Size").Value = $($_.Length)
            $fileAttachment.Item($classFA, "AddedDate").Value = [DateTime]::Now.ToUniversalTime()
            $fileAttachment.Item($classFA, "Content").Value = $($fileStream)
            
            $projectionType = Get-SCSMTypeProjection -Name System.WorkItem.ServiceRequestProjection$
            $projection = Get-SCSMObjectProjection -ProjectionName $projectionType.Name -Filter "ID -eq $($translatedFolderName)"

            $projection.__base.Add($fileAttachment, $fileAttachmentRel.Target)
            $projection.__base.Commit()

            $actionLogGuid = [Guid]::NewGuid().ToString()
            $managementPack = Get-SCManagementPack -Name "System.WorkItem.Library"
            $actionType = "System.WorkItem.ActionLogEnum.FileAttached"

            $actionLogEntry = New-Object Microsoft.EnterpriseManagement.Common.CreatableEnterpriseManagementObject($managementGroup, $classAL)

            $actionLogEntry.Item($classAL, "Id").Value = $actionLogGuid
            $actionLogEntry.Item($classAL, "DisplayName").Value = $actionLogGuid
            $actionLogEntry.Item($classAL, "ActionType").Value = $managementPack.GetEnumerations().GetItem($actionType)
            $actionLogEntry.Item($classAL, "Title").Value = "Attached File (Migrated)"
            $actionLogEntry.Item($classAL, "EnteredBy").Value = "SYSTEM"
            $actionLogEntry.Item($classAL, "Description").Value = $($_.Name)
            $actionLogEntry.Item($classAL, "EnteredDate").Value = [System.DateTime]::UtcNow

            $projection.__base.Add($actionLogEntry, $actionLogRel.Target)
            $projection.__base.Commit()

            $fileStream.Close();

        }
    }
}

function Get-ConvertedServiceRequestNumber([string]$oldSR, [string]$diffPath)
{
    $diffTable = Import-Csv $diffPath
    $newSR = $diffTable | Where-Object {$_.previousID -eq $($oldSR)} | ForEach-Object CurrentID
    return $newSR
}

function Get-ServiceRequestGuid([string]$serviceRequestID)
{
    $class = Get-SCSMClass System.WorkItem.ServiceRequest$
    [Guid]$serviceRequestGuid = Get-SCSMObject -Class $class | Where-Object {$_.Id -eq $($serviceRequestID)} | ForEach-Object __internalID
    return $serviceRequestGuid
}

function Get-Guid([string]$serviceRequestID)
{
    $class = Get-SCSMClass System.WorkItem.Activity.ManualActivity$
    [Guid]$serviceRequestGuid = Get-SCSMObject -Class $class | Where-Object {$_.Id -eq $($serviceRequestID)} | ForEach-Object __internalID
    return $serviceRequestGuid
}

function Get-ParallelActivityGuid([string]$parallelActivityID)
{
    $class = Get-SCSMClass System.WorkItem.Activity.ParallelActivity$
    [Guid]$parallelActivityGuid = Get-SCSMObject -Class $class | Where-Object {$_.Id -eq $($parallelActivityID)} | ForEach-Object __internalID
    return $parallelActivityGuid
}

Import-SMLets
#Test-EnumSanity $sourceServiceRequestCsv #test todo
Get-Down
Import-ServiceRequest $sourceServiceRequestCsv $diffTableCsv $sourceReviewActivityCsv $sourceManualActivityCsv $sourceParallelActivityCsv
Import-ServiceRequestActivityLog $sourceActivityLogCsv $diffTableCsv
Import-ServiceRequestFileAttachments $sourceFileAttachmentsDir $diffTableCsv