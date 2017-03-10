<#
.SYNOPSIS
    Imports Incidents to System Center Service Manager from a CSV

.DESCRIPTION
    Imports Incidents and their useful/related data from a CSV file (probably created by the included Export-Incident.ps1 script)

    The purpose of this script is to help automate the process of migrating to a new Service Manager environment, bringing
    only the data you want. In this script's current life, we import all incidents that are contained in your CSV. We require
    the following:
        - A CSV containing all the incidents and their supporting data
        - A CSV containing the incidents' activity log data
        - A "surrogate" user in the event that an assigned/affected user no longer exists
        - A directory containing all incident attachments
        - Write access to a directory to create a temporary "diffTable" CSV. By default, this script recreates all incidents using
        the same IR numbers as before. If this is a lab (or a production environment that you purposefully want to start with new
        numbers), this CSV is used to keep track of how the new IR number is related to it's previous (old) IR number

    Before importing data, we run a "sanity check" against your list enums. If your incident lists don't contain everything your
    previous environment had, we'll terminate early. Therefore, we recommend that you import any required Management Packs prior
    to running this script

.PARAMETER sourceIncidentCsv
    The path and name of a CSV file that contains all Incident data
    
    This field is required, and should contain the path to the CSV

.PARAMETER surrogateAffectedUser
    A user (in "Display Name" format) that will be assigned the "Affected User" of any incidents that contain an affected user
    who no longer exists

    This field is required. If you don't tell Service Manager to sync with AD first, all of your Incidents will likely have their
    affected user set to this person

.PARAMETER surrogateAssignedTo
    A user (in "Display Name" format) that will be assigned the "Assigned To" of any incidents that contain an AssignedTo user
    who no longer exists

    This field is required. If you don't tell Service Manager to sync with AD first, all of your Incidents will likely be assigned
    to this person. If this is a big import, you should get them a coffee

.PARAMETER sourceFileAttachmentsDir
    The path of a directory that contains Incident attachments

    This field is not required, and should be a location that you have read access to. This directory should contain a list of directories
    named by ID (IR1, etc.). We expect to find the IR's attachments in these subdirectories

    If this parameter is not specified, we don't attempt to import attachments

.PARAMETER sourceActivityLogCsv
    The path and name of a CSV file that contains all Incident Activity Log data
    
    This field is required, and should contain the path to the CSV

.PARAMETER diffTableCsv
    The path and name of a CSV file that the script will write out to keep track of existing and new incident ID information
    
    This field is required, and should contain the path to the CSV
    This path needs to be writable by the account running this script

.EXAMPLE
    Import all Incidents from a Service Manager environment

    Import-Incident -sourceIncidentCsv .\incidents.csv -surrogateAffectedUser "Douglas, Fred" -surrogateAssignedTo "Douglas, Fred" -sourceFileAttachmentsDir .\exported_IR_attachments -sourceActivityLogCsv .\incidents_activitylogs.csv -diffTable .\difftable.csv

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
    [Parameter(Mandatory = $True, HelpMessage = "The CSV containing the incidents you want to import")]
    [string]$sourceIncidentCsv,

    [Parameter(Mandatory = $True, HelpMessage = "The user in DisplayName format that will be used if the affected user no longer exists")]
    [string]$surrogateAffectedUser,

    [Parameter(Mandatory = $True, HelpMessage = "The user in DisplayName format that will be used if the assigned to user no longer exists")]
    [string]$surrogateAssignedTo,

    [Parameter(Mandatory = $True, HelpMessage = "Directory containing previously-exported attachments. Will skip attachments if not specified")]
    [string]$sourceFileAttachmentsDir,

    [Parameter(Mandatory = $True, HelpMessage = "CSV containing previously-exported activity log entries. Will skip attachments if not specified")]
    [string]$sourceActivityLogCsv,

    [Parameter(Mandatory = $True, HelpMessage = "This script needs to write out a file that lists your old IR numbers alongside your new ones. Enter the full path here.")]
    [string]$diffTableCsv
)

$irDiffTableArray=@()

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
    for ($i=$magic; $i -gt 1; $i--) {

        Write-Progress -Activity "Take a deep breath and ready your wand--its almost time for some magic!" -Status "Reticulating splines" -SecondsRemaining $i
        Start-Sleep 1
    }
}

function Test-EnumSanity([string]$filePath)
{
    $incidentClass = Get-SCSMClass -Name System.WorkItem.Incident$
    $userClass = Get-SCSMClass -Name System.Domain.User$  
    $incidentAffectedUserRelClass = Get-SCSMRelationshipClass -Name System.WorkItemAffectedUser$
    $assignedToUserRelClass  = Get-SCSMRelationshipClass -Name System.WorkItemAssignedToUser$ 
    $stopCondition = $false
    $i = 0
    $totalCount = (Import-Csv $filePath).Count

    Import-Csv $filePath | ForEach-Object {
        Write-Progress -Activity "Verifying environment Enum values" -Status "Checking existing incident $($_.Id)" -PercentComplete (($i / $totalCount) * 100)
        
        $testSourceDisplayName = $_.SourceDisplayName
        $testStatusDisplayName = $_.StatusDisplayName
        $testTierQueueDisplayName = $_.TierQueueDisplayName
        $testLastModifiedSource = $_.LastModifiedSource
        $testClassificationDisplayName = $_.ClassificationDisplayName
        $testResolutionCategoryDisplayName = $_.ResolutionCategoryDisplayName
        $testImpact = $_.Impact
        $testUrgency = $_.Urgency

        if (((Get-SCSMEnumeration | Where-Object { $_.DisplayName -eq $testSourceDisplayName }).Count -eq 0) -xor ($testSourceDisplayName -eq ""))
        {
            Write-Error -Message "The enum value $testSourceDisplayName for Source does not exist in this environment." -ErrorAction Continue
            $stopCondition = $true
        }
        if (((Get-SCSMEnumeration | Where-Object { $_.DisplayName -eq $testStatusDisplayName }).Count -eq 0) -xor ($testStatusDisplayName -eq ""))
        {
            Write-Error -Message "The enum value $testStatusDisplayName for Status does not exist in this environment." -ErrorAction Continue
            $stopCondition = $true
        }
        if (((Get-SCSMEnumeration | Where-Object { $_.DisplayName -eq $testTierQueueDisplayName }).Count -eq 0) -xor ($testTierQueueDisplayName -eq ""))
        {
            Write-Error -Message "The enum value $testTierQueueDisplayName for TierQueue does not exist in this environment." -ErrorAction Continue
            $stopCondition = $true
        }
        #if ((Get-SCSMEnumeration | Where-Object { $_.DisplayName -eq $testLastModifiedSource }).Count -eq 0)
        #{
        #    Write-Error -Message "The enum value $testLastModifiedSource for LastModifiedSource does not exist in this environment." -ErrorAction Continue
        #    $stopCondition = $true
        #}
        if (((Get-SCSMEnumeration | Where-Object { $_.DisplayName -eq $testClassificationDisplayName }).Count -eq 0) -xor ($testClassificationDisplayName -eq ""))
        {
            Write-Error -Message "The enum value $testClassificationDisplayName for Classification does not exist in this environment." -ErrorAction Continue
            $stopCondition = $true
        }
        if (((Get-SCSMEnumeration | Where-Object { $_.DisplayName -eq $testResolutionCategoryDisplayName }).Count -eq 0) -xor ($testResolutionCategoryDisplayName -eq ""))
        {
            Write-Error -Message "The enum value $testResolutionCategoryDisplayName for ResolutionCategory does not exist in this environment." -ErrorAction Continue
            $stopCondition = $true
        }
        #if ((Get-SCSMEnumeration | Where-Object { $_.DisplayName -eq $testImpact }).Count -eq 0)
        #{
        #    Write-Error -Message "The enum value $testImpact for Impact does not exist in this environment." -ErrorAction Continue
        #    $stopCondition = $true
        #}
        #if ((Get-SCSMEnumeration | Where-Object { $_.DisplayName -eq $testUrgency }).Count -eq 0)
        #{
        #    Write-Error -Message "The enum value $testUrgency for Urgency does not exist in this environment." -ErrorAction Continue
        #    $stopCondition = $true
        #}
        $i++
    }

    if ($stopCondition -eq $true)
    {
        Write-Error "Enum value dependency failure. Please add missing values to lists or import missing Management Packs." -ErrorAction Stop
    }
}

function Import-Incident([string]$filePath, [string]$diffPath)
{
    $incidentClass = Get-SCSMClass -Name System.WorkItem.Incident$
    $userClass = Get-SCSMClass -Name System.Domain.User$  
    $incidentAffectedUserRelClass = Get-SCSMRelationshipClass -Name System.WorkItemAffectedUser$
    $assignedToUserRelClass  = Get-SCSMRelationshipClass -Name System.WorkItemAssignedToUser$ 
    $i = 0
    $totalCount = (Import-Csv $filePath).Count

    Import-Csv $filePath | ForEach-Object {
        $irObject = New-Object PSObject
        $irObject | Add-Member -MemberType NoteProperty -Name "PreviousId" -Value $_.Id
        $affectedUser = ""
        $assignedTo = ""

        $irId = $_.Id
        $irSourceDisplayName = $_.SourceDisplayName
        $irStatusDisplayName = $_.StatusDisplayName
        $irResolutionDescription = $_.ResolutionDescription
        $irTierQueueDisplayName = $_.TierQueueDisplayName
        $irLastModifiedSource = $_.LastModifiedSource
        $irClassificationDisplayName = $_.ClassificationDisplayName
        $irResolutionCategoryDisplayName = $_.ResolutionCategoryDisplayName
        $irPriority = $_.Priority
        $irImpact = $_.Impact
        $irUrgency = $_.Urgency
        $irTitle = $_.Title
        $irDescription = $_.Description
        $irContactMethod = $_.ContactMethod
        $irUserInput = $_.UserInput

        #$irPlannedCost = $_.PlannedCost
        #$irActualCost = $_.ActualCost
        #$irPlannedWork = $_.PlannedWork
        #$irActualWork = $_.ActualWork
        
        if ($irUserInput.Length -eq 0) {
            $irUserInput = $null
        }

        #Change the Id below to "IR{0}" to generate a new IR instead of reusing the same IR number
        $incidentHashTable = @{
            Source = $irSourceDisplayName;
            Status = $irStatusDisplayName;
            ResolutionDescription = $irResolutionDescription;
            TierQueue = $irTierQueueDisplayName;
            LastModifiedSource = $irLastModifiedSource;
            Classification = $irClassificationDisplayName;
            ResolutionCategory = $irResolutionCategoryDisplayName;
            Priority = $irPriority;
            Impact = $irImpact;
            Urgency = $irUrgency;
            Id = $irId;
            Title = $irTitle;
            Description = $irDescription;
            UserInput = $irUserInput;
        }
        try {
            [datetime]$irTargetResolutionTime = $_.TargetResolutionTime
            $incidentHashTable += @{TargetResolutionTime = $irTargetResolutionTime}
        } catch { }
        try {
            [bool]$irEscalated = $_.Escalated
            if($irEscalated -eq "True") {$irEscalated = $true} else {$irEscalated = $false}
            $incidentHashTable +=@{Escalated = $irEscalated}
        }
        catch { }
        try {
            [bool]$irNeedsKnowledgeArticle = $_.NeedsKnowledgeArticle
            if($irNeedsKnowledgeArticle -eq "True") {$irNeedsKnowledgeArticle = $true} else {$irNeedsKnowledgeArticle = $false}
            $incidentHashTable +=@{NeedsKnowledgeArticle = $irNeedsKnowledgeArticle}
        }
        catch { }
        try {
            [bool]$irHasCreatedKnowledgeArticle = $_.HasCreatedKnowledgeArticle
            if($irHasCreatedKnowledgeArticle -eq "True") {$irHasCreatedKnowledgeArticle = $true} else {$irHasCreatedKnowledgeArticle = $false}
            $incidentHashTable +=@{HasCreatedKnowledgeArticle = $irHasCreatedKnowledgeArticle}
        }
        catch { }
        try {
            [datetime]$irClosedDate = $_.ClosedDate
            $incidentHashTable +=@{ClosedDate = $irClosedDate}
        }
        catch { }
        try {
            [datetime]$irResolvedDate = $_.ResolvedDate
            $incidentHashTable +=@{ResolvedDate = $irResolvedDate}
        }
        catch { }
        try {
            [datetime]$irScheduledStartDate = $_.ScheduledStartDate
            $incidentHashTable +=@{ScheduledStartDate = $irScheduledStartDate}
        }
        catch { }
        try {
            [datetime]$irScheduledEndDate = $_.ScheduledEndDate
            $incidentHashTable +=@{ScheduledEndDate = $irScheduledEndDate}
        }
        catch { }
        try {
            [datetime]$irActualStartDate = $_.ActualStartDate
            $incidentHashTable +=@{ActualStartDate = $irActualStartDate}
        }
        catch { }
        try {
            [datetime]$irActualEndDate = $_.ActualEndDate
            $incidentHashTable +=@{ActualEndDate = $irActualEndDate}
        }
        catch { }
        try {
            [bool]$irIsDowntime = $_.IsDowntime
            if($irIsDowntime -eq "True") {$irIsDowntime = $true} else {$irIsDowntime = $false}
            $incidentHashTable +=@{IsDowntime = $irIsDowntime}
        }
        catch { }
        try {
            [bool]$irIsParent = $_.IsParent
            if($irIsParent -eq "True") {$irIsParent = $true} else {$irIsParent = $false}
            $incidentHashTable +=@{IsParent = $irIsParent}
        }
        catch { }
        try {
            [datetime]$irScheduledDowntimeStartDate = $_.ScheduledDowntimeStartDate
            $incidentHashTable +=@{ScheduledDowntimeStartDate = $irScheduledDowntimeStartDate}
        }
        catch { }
        try {
            [datetime]$irScheduledDowntimeEndDate = $_.ScheduledDowntimeEndDate
            $incidentHashTable +=@{ScheduledDowntimeEndDate = $irScheduledDowntimeEndDate}
        }
        catch { }
        try {
            [datetime]$irActualDowntimeStartDate = $_.ActualDowntimeStartDate
            $incidentHashTable +=@{ActualDowntimeStartDate = $irActualDowntimeStartDate}
        }
        catch { }
        try {
            [datetime]$irActualDowntimeEndDate = $_.ActualDowntimeEndDate
            $incidentHashTable +=@{ActualDowntimeEndDate = $irActualDowntimeEndDate}
        }
        catch { }
        try {
            [datetime]$irRequiredBy = $_.RequiredBy
            $incidentHashTable +=@{RequiredBy = $irRequiredBy;}
        }
        catch { }
        try {
            [datetime]$irCreatedDate = $_.CreatedDate
            $incidentHashTable +=@{CreatedDate = $irCreatedDate}
        }
        catch { }
        try {
            [datetime]$irFirstAssignedDate = $_.FirstAssignedDate
            $incidentHashTable +=@{FirstAssignedDate = $irFirstAssignedDate}
        }
        catch { }
        try {
            [datetime]$irFirstResponseDate = $_.FirstResponseDate
            $incidentHashTable +=@{FirstResponseDate = $irFirstResponseDate}
        }
        catch { }
        #try {
        #    [datetime]$irTimeAdded = $_.TimeAdded
        #    $incidentHashTable +=@{TimeAdded = $irTimeAdded}
        #}
        #catch { }

        $surrogateAffectedUser = "insert name here for now"
        $surrogateAssignedTo = "insert name here for now"

        try {
            $affectedUser = Get-SCSMObject -Class $UserClass -Filter "DisplayName -eq $($_.AffectedUser)"
        } catch {
            $affectedUser = Get-SCSMObject -Class $UserClass -Filter "DisplayName -eq $($surrogateAffectedUser)"
            Write-Information "Used a surrogate Affected User for $newIncidentId"
        }

        try {
            $assignedTo = Get-SCSMObject -Class $UserClass -Filter "DisplayName -eq $($_.AssignedTo)"
        } catch {
            $assignedTo = Get-SCSMObject -Class $UserClass -Filter "DisplayName -eq $($surrogateAssignedTo)"
            Write-Information "Used a surrogate AssignedTo User for $newIncidentId"
        }

        $newIncident = New-SCSMObject -Class $incidentClass -PropertyHashtable $incidentHashTable -PassThru
        $newIncidentId = $newIncident.Id
        $newIncidentGuid = $newIncident.__internalID

        $irObject | Add-Member -MemberType NoteProperty -Name "CurrentId" -Value $newIncidentId
        $irObject | Add-Member -MemberType NoteProperty -Name "CurrentGuid" -Value $newIncidentGuid
        $irDiffTableArray += $irObject

        if (($affectedUser).Count -eq 1) { New-SCSMRelationshipObject -Relationship $incidentAffectedUserRelClass -Source $newIncident -Target $affectedUser -Bulk }
        if (($affectedUser).Count -gt 1) { Write-Information "Unable to add the Affected User to $newIncidentId because there was more than one for some reason." }
        if (($assignedTo).Count -eq 1) { New-SCSMRelationshipObject -Relationship $assignedToUserRelClass -Source $newIncident -Target $assignedTo -Bulk }
        if (($assignedTo).Count -gt 1) { Write-Information "Unable to add the AssignedTo User to $newIncidentId because there was more than one for some reason." }
        Write-Progress -Activity "Importing Incidents" -Status "Re-created new incident $newIncidentId from previous incident $($_.Id)" -PercentComplete (($i / $totalCount) * 100)
        $i++;
    }

    $irDiffTableArray | Export-Csv $diffPath
}

function Import-IncidentActivityLog([string]$filePath, [string]$diffPath)
{
    $i = 0
    $totalCount = (Import-Csv $filePath).Count
    $theDiffList = Import-Csv $diffPath 

    Import-Csv $filePath | ForEach-Object {
        $newGUID = ([Guid]::NewGuid()).ToString()
        $oldIR = $_.RelatedIncident
        $newIR = $theDiffList | Where-Object {$_.PreviousID -eq $($oldIR)} | ForEach-Object CurrentId
        $smObject = Get-SCSMObject ($theDiffList | Where-Object {$_.PreviousID -eq $($oldIR)} | ForEach-Object CurrentGuid)
        $newComment = $_.Comment
        $newEnteredBy = $_.EnteredBy
        $newEnteredDate = $_.EnteredDate
        switch($_.LogType) {
            "AnalystComment" {
                if($_.isPrivate -eq "True") {[bool]$isPrivate = $true} else {[bool]$isPrivate = $false}
                $commentProjection = @{__CLASS = "System.WorkItem.Incident";
                                       __SEED = $smObject;
                AnalystComments = @{__CLASS = "System.WorkItem.TroubleTicket.AnalystCommentLog";
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
                $commentProjection = @{__CLASS = "System.WorkItem.Incident";
                                       __SEED = $smObject;
                UserComments = @{__CLASS = "System.WorkItem.TroubleTicket.UserCommentLog";
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
        
        New-SCSMObjectProjection -Type "System.WorkItem.IncidentPortalProjection" -Projection $commentProjection
        Write-Progress -Activity "Importing Incident Activity Log" -Status "Migrated activity log for $newIR" -PercentComplete (($i / $totalCount) * 100)
        $i++;
    }
}

function Import-IncidentFileAttachments([string]$filePath, [string]$diffPath)
{
    $attachmentList = Get-ChildItem -Directory $filePath
    $i = 0
    $totalCount = $attachmentList.Count

    Get-ChildItem -Directory $filePath | ForEach-Object {
        $translatedIncidentNumber = Get-ConvertedIncidentNumber $_.Name $diffPath
        Get-ChildItem -File ($_.FullName) | ForEach-Object {
            Write-Verbose "Adding $($_.FullName) to $translatedIncidentNumber"
            Set-SCSMIncident -ID $translatedIncidentNumber -AttachmentPath $_.FullName -Verbose
        }
        Write-Progress -Activity "Importing Incident Attachments" -Status "Added attachments to $translatedIncidentNumber" -PercentComplete (($i / $totalCount) * 100)
        $i++;
    }
}

function Get-ConvertedIncidentNumber([string]$oldIR, [string]$diffPath)
{
    $diffTable = Import-Csv $diffPath
    $newIR = $diffTable | Where-Object {$_.previousID -eq $($oldIR)} | ForEach-Object CurrentID
    return $newIR
}

function Get-IncidentGuid([string]$incidentID)
{
    $class = Get-SCSMClass System.WorkItem.Incident$
    [Guid]$incidentGuid = Get-SCSMObject -Class $class | Where-Object {$_.Id -eq $($incidentID)} | ForEach-Object __internalID
    return $incidentGuid
}

Import-SMLets
Test-EnumSanity $sourceIncidentCsv
Get-Down
Import-Incident $sourceIncidentCsv $diffTableCsv
Import-IncidentActivityLog $sourceActivityLogCsv $diffTableCsv
Import-IncidentFileAttachments $sourceFileAttachmentsDir $diffTableCsv