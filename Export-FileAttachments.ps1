<#
.SYNOPSIS
    Exports file attachments from work items or configuration items

.DESCRIPTION
    Ingests a Work Item or Configuration Item GUID and exports all file attachments to a unique folder named after the object's display name
    Requires the SMlets PowerShell module

.PARAMETER smObjectGUID
    Internal GUID of a Work Item or Configuration Item

.PARAMETER attachmentPath
    The location where the exported file attachments will be saved

.EXAMPLE
    Export attachments for a single Incident

        .\Export-FileAttachments.ps1 -smObjectGUID eaa86412-bd85-1f99-b383-7f519063d8a8 -attachmentPath "C:\scsm\incident_attachments"

    Export attachments for all Incidents

        $class = Get-SCSMClass System.WorkItem.Incident$
        $incidents = Get-SCSMObject -Class $class
        foreach ($incident in $incidents)
        {
            .\Export-FileAttachments.ps1 -smObjectGUID $incident.Get_Id() -attachmentPath "C:\scsm\incident_attachments"  
        }

.NOTES
    Jacob Thornberry
    Adapted from Patrik Sundqvist, http://blogs.litware.se/?p=1369
    @jakertberry
    Made with ❤ in MKE
#>

[CmdletBinding()]
Param (
    [Parameter(Mandatory = $True, HelpMessage = "The internal ID of a work item or configuration item")]
    [Guid]$smObjectGUID,

    [Parameter(Mandatory = $True, HelpMessage = "Location where file attachments should be exported")]
    [string]$attachmentPath
)

function Test-OutputPath([string]$filePath)
{
    if (!(Test-Path $filePath))
    {
        Write-Error "Unable to export attachments, the path $filePath does not exist" -ErrorAction Stop
    }
} 
function Import-SMLets
{
    if (!(Get-Module SMLets))
    {
        Write-Progress -Activity "Initializing" -Status "Importing the SMLets Module"
        Import-Module SMLets -Force -ErrorAction Stop
    }
}

function Export-FileAttachments([Guid]$objectID, [string]$filePath)
{
    $filePath = $filePath.TrimEnd("\")
    $i = 0

    $scsmObject = Get-SCSMObject -Id $objectID
    $scsmObjectID = $scsmObject.Id

    [Guid]$workItemAttachment = "aa8c26dc-3a12-5f88-d9c7-753e5a8a55b4"
    [Guid]$configurationItemAttachment = "095ebf2a-ee83-b956-7176-ab09eded6784"

    $workItemAttachmentClass = Get-SCSMRelationshipClass -Id $workItemAttachment
    $workItemClass = Get-SCSMClass -Name System.WorkItem$
    
    $configurationItemAttachmentClass = Get-SCSMRelationshipClass -Id $configurationItemAttachment
    $configurationItemClass = Get-SCSMClass -Name System.ConfigItem$
    
    if ($scsmObject.IsInstanceOf($workItemClass)) 
    {
        Write-Verbose "$scsmObjectID is a Work Item"
        $fileAttachments = Get-SCSMRelatedObject -SMObject $scsmObject -Relationship $workItemAttachmentClass
    } 
    elseif ($scsmObject.IsInstanceOf($configurationItemClass)) 
    {
        Write-Verbose "$scsmObjectID is a Configuration Item"
        $fileAttachments = Get-SCSMRelatedObject -SMObject $scsmObject -Relationship $configurationItemAttachmentClass
    }
    else
    {
        Write-Error "Class type is unknown or unsupported for $scsmObjectID. Only WorkItems and ConfigItems are supported" -ErrorAction Stop
    }

    if ($fileAttachments -ne $Null)
    {
        $fileAttachmentPath = $filePath + "\" + $scsmObject.Id
        New-Item -Path ($fileAttachmentPath) -ItemType Directory -Force | Out-Null
        $totalAttachments = $fileAttachments.Count
        $fileAttachments | ForEach-Object {
            Write-Progress -Activity "Exporting Attachments in $scsmObjectID" -Status $_.DisplayName -PercentComplete (($i / $totalAttachments) * 100)
            try {
                $exportedFile = [IO.File]::OpenWrite(($fileAttachmentPath + "\" + $_.DisplayName))
                $memoryStream = New-Object IO.MemoryStream
                $buffer = New-Object byte[] 8192
                [int]$bytesRead | Out-Null
                while (($bytesRead = $_.Content.Read($buffer, 0, $buffer.Length)) -gt 0)
                    {
                        $memoryStream.Write($buffer, 0, $bytesRead)
                    }        
                $memoryStream.WriteTo($exportedFile)
            } catch {
                $errorMessage = $_.Exception.Message
                $failedItem = $_.Exception.ItemName
                Write-Error "Failed to export file $failedItem with error $errorMessage" 
            } finally {
                $exportedFile.Close()
                $memoryStream.Close()
            } 
            $i++
        }
    }

}

Import-SMLets
Test-OutputPath $attachmentPath
Export-FileAttachments $smObjectGUID $attachmentPath