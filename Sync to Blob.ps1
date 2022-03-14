#requires -version 3

<#
.SYNOPSIS
    Watch a local folder and copy changed files to Azure blob storage with option to copy files before starting watching the folder
.PARAMETER folder
    Folder to monitor
.PARAMETER ResourceGroupName
    Azure resource group where the storage account resides
.PARAMETER StorageAccountName
    Name of the Azure storage account where the blob container resides
.PARAMETER ContainerName
    Name of the Azure blob storage container
.PARAMETER filter
    Only sync files matching this file pattern
.PARAMETER tenantId
    The Azure tenant id
.PARAMETER credential
    PS credential for a Service Principal to authenticate to Azure to be written to the credentials file
.PARAMETER credentialFile
    Credential file to read PS credential from or write to
.PARAMETER BlobType
    The type of blob
.PARAMETER minutesToRun
    How long in minutes to run the sync for. Specifying 0 will run inifinitely
.PARAMETER minutesBack
    How far back in minutes to look for locally changed files to sync to Azure before starting the folder watcher
.PARAMETER tier
    The storage tier to set for the uploaded files
.PARAMETER tags
    Tags to set for the uploaded files
.PARAMETER includeSubdirectories
    Watch for changed files in sub folders and sync those to corresponding folders in Azure
.PARAMETER noWatch
    Perform the initial file sync but do not watch the folder for changed files
.PARAMETER millisecondsApart
    If changes for the file are less than this time apart, consider them to be part of the same change and do not perform further syncs
.PARAMETER existingOnly
    Only sync files if they already exist in Azure
.PARAMETER force
    Overwrite files in Azure even if local copy is older or overwrite credentials file if it already exists
.EXAMPLE
    & '.\Sync to Blob.ps1' -folder 'C:\ARM Templates' -ContainerName templates -StorageAccountName slartibartfast -ResourceGroupName Zaphod-RG -existingOnly -force -includeSubdirectories -minutesback 60

    Sync any files changed in the 'C:\ARM Templates' folder or sub-folders in the last 60 minutes to the given blob storage in the storage account and resource group specified if that file already exists in Azure.
        Once this sync finishes, watch the 'C:\ARM Templates' folder and sub-folders for changed files and sync them to Azure. The watcher will run until interrutped/stopped
.EXAMPLE
    & '.\Sync to Blob.ps1' -credential $credential -credentialfile c:\creds\GuysAzure.xml -tenantId f886e631-a11a-4f8b-89e3-10dff29dc15c

    Store the specified tenant id and the application id and secret contained in the $credential variable (create interactively with Get-Credential) to the credential file c:\creds\GuysAzure.xml
        This allows the script to be used non-interactively such as via a scheduled task
.EXAMPLE
    & '.\Sync to Blob.ps1' -folder 'C:\ARM Templates' -ContainerName templates -StorageAccountName slartibartfast -ResourceGroupName Zaphod-RG -credentialfile c:\creds\GuysAzure.xml -minutesToRun 60 -filter *.json

    Authenticate to Azure using the service principal credentials stored in the file c:\creds\GuysAzure.xml and then watch the 'C:\ARM Templates' folder for changed or new .json files for 60 minutes and sync them to Azure
.NOTES
    Locally deleted files will not be deleted in Azure

    Modification History:

    2022/03/11  @guyrleech  Initial version
#>

<#
Copyright © 2022 Guy Leech

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction,
including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

[CmdletBinding(DefaultParameterSetName='Watch')]

Param
(
    [Parameter(Mandatory=$true,ParameterSetName='Watch')]
    [string]$folder ,
    [Parameter(Mandatory=$true,ParameterSetName='Watch')]
    [ValidateNotNullOrEmpty()]
    [string]$ResourceGroupName ,
    [Parameter(Mandatory=$true,ParameterSetName='Watch')]
    [ValidateNotNullOrEmpty()]
    [string]$StorageAccountName ,
    [Parameter(Mandatory=$true,ParameterSetName='Watch')]
    [ValidateNotNullOrEmpty()]
    [string]$ContainerName,
    [Parameter(Mandatory=$false,ParameterSetName='Watch')]
    [string]$filter = '*' ,
    [string]$tenantId ,
    [Parameter(Mandatory=$true,ParameterSetName='Save')]
    [PSCredential]$credential ,
    [Parameter(Mandatory=$true,ParameterSetName='Save')]
    [Parameter(ParameterSetName='Watch')]
    [string]$credentialFile ,
    [Parameter(Mandatory=$false,ParameterSetName='Watch')]
    [ValidateSet('Append','Block','Page')]
    [string]$BlobType = 'Block' ,
    [Parameter(Mandatory=$false,ParameterSetName='Watch')]
    [decimal]$minutesToRun ,
    [Parameter(Mandatory=$false,ParameterSetName='Watch')]
    [decimal]$minutesBack ,
    [Parameter(Mandatory=$false,ParameterSetName='Watch')]
    [ValidateSet('Hot', 'Cool', 'Archive')]
    [string]$tier ,
    [Parameter(Mandatory=$false,ParameterSetName='Watch')]
    [hashtable]$tags ,
    [Parameter(Mandatory=$false,ParameterSetName='Watch')]
    [switch]$includeSubdirectories ,
    [Parameter(Mandatory=$false,ParameterSetName='Watch')]
    [switch]$noWatch ,
    [Parameter(Mandatory=$false,ParameterSetName='Watch')]
    [int]$millisecondsApart = 100 ,
    [Parameter(Mandatory=$false,ParameterSetName='Watch')]
    [switch]$existingOnly ,
    [switch]$force
)

## Can't use [System.Web.MimeMapping]::GetMimeMapping() in pwsh 7
Function Get-MimeMapping
{
    Param
    (
        [Parameter(Mandatory=$true)]
        [string]$file
    )

    [string]$result = 'application/octet-stream'

    if ( -Not ( Get-PSDrive -name HKCR -ErrorAction SilentlyContinue ) )
    {
        $null = New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT
    }

    if( $mimetype = ( Get-ItemProperty -Path "HKCR:\$([System.IO.Path]::GetExtension( $file ))" -Name 'Content Type' -ErrorAction SilentlyContinue | Select-object -ExpandProperty 'Content Type') )
    {
        $result = $mimetype
    }
    else
    {
        Write-Warning -Message "Unable to find HKCR content type for $file"
    }

    $result # return
}

if( $PsCmdlet.ParameterSetName -eq 'Save' )
{
    if( [string]::ISNullOrEmpty( $credentialFile ))
    {
        Throw "Must specify credential file name to write to via -credentialFile"
    }
    elseif( -Not $credential )
    {
        Throw "Must specify credential to write to file via -credential"
    }
    elseif( (Test-Path -path $credentialFile -ErrorAction SilentlyContinue )-and -Not $force )
    {
        Throw "Credential file already exists - use -force to overwrite"
    }
    else
    {
        ## if tenant id is specified we'll add this too otherwise we'll need it passed when we use the credentials
        [hashtable]$export = @{ 'Credentials' = $credential }
        if( -Not [string]::ISNullOrEmpty( $tenantId ))
        {
            $export.Add( 'TenantId' , $tenantId )
        }
        else
        {
            Write-Warning -Message "Tenant id not specified so must be specified when using this credentials file"
        }
        $export | Export-CliXML -Path $credentialFile -Force:$force -Depth 2
        if( -Not $? )
        {
            Throw "Problem writing credentials to $credentialFile"
        }
    }
}
else
{
    Import-Module -Name Az.Storage -ErrorAction Stop -Verbose:$false

    if( -Not [string]::ISNullOrEmpty( $credentialFile ))
    {
        if( -Not ( $savedCredentials = Import-cliXML -Path $credentialFile ) )
        {
            Throw "Failed to import credentials from $credentialFile"
        }
        if( [string]::ISNullOrEmpty( $tenantId ))
        {
            if( -Not ( $tenantId = $savedCredentials[ 'tenantId']) )
            {
                Throw "Tenant id not specified and not stored in credential file"
            }
        }
        $connection = $null
        $connection = Connect-AzAccount -Credential $savedCredentials.Credentials -ServicePrincipal -Tenant $tenantId
        if( -Not $connection )
        {
            Throw "Failed to connect to tenant id $tenantId with service principal"
        }
        else
        {
            Write-Verbose -Message "Connected to subscription `"$($connection.context.Subscription.Name )`""
        }
    }

    $startTime = [datetime]::Now

    [int]$basePathLength = (Resolve-Path -Path $folder).Path.Length + 1 ## + 1 to get passed the \

    $storageAccountParameters = @{
        'ResourceGroupName' = $ResourceGroupName
        'Name' = $StorageAccountName
    }

    $storageContainerParameters = @{
        'Container' = $ContainerName
    }

    ## https://adamtheautomator.com/copy-files-to-azure-blob-storage/

    if( -Not ( $storageAccount = Get-AzStorageAccount @storageAccountParameters ) )
    {
        Throw "Failed to get storage account $($storageAccountParamters['Name']) in resource group $($storageAccountParameters['ResourceGroupName'])"
    }

    if( -Not ( $storageContainer = $storageAccount | Get-AzStorageContainer @storageContainerParameters ) )
    {
        Throw "Failed to get storage container $($storageContainerParameters['Container']) in resource group $($storageAccountParameters['ResourceGroupName'])"
    }

    [hashtable]$storageBlobContentParameters = @{
        'File' = $null
        ## need to include any sub folders so they get created in Azure so we strip off the path before the root of the folder we're monitoring
        'Blob' =  $null
        'BlobType' = $BlobType
        'Force'= $force
        'Properties' = @{
            "ContentType" = $Null
        }
    }
    if( $PSBoundParameters[ 'tier'])
    {
        $storageBlobContentParameters.Add( 'StandardBlobTier' , $tier)
    }
    if( $PSBoundParameters[ 'tags'])
    {
        $storageBlobContentParameters.Add( 'Tag' , $tags)
    }

    [hashtable]$existingItems = @{}
    $storageContainer | Get-AzStorageBlob | ForEach-Object `
    {
        $existingItems.Add( $_.Name , $_.LastModified.UtcDateTime )
    }
    Write-Verbose -Message "Got $($existingItems.Count) items from blob"

    if( $PSBoundParameters.ContainsKey( 'minutesBack'))
    {
        [datetime]$fromTime = $startTime.AddMinutes( -$minutesBack )
        if( $minutesBack -eq 0 ) ## special value meaning include everything
        {
            $fromTime = New-Object -typename datetime
        }
        Write-Verbose -Message "Looking for files changed since $(Get-Date -Date $fromTime -Format G)"
        ForEach( $file in (Get-ChildItem -Path $folder -File -Recurse:$includeSubdirectories -Force | Where-Object LastWriteTimeUTC -ge $fromTime ))
        {
            [bool]$sync = $true
            [string]$itemPath = $file.fullname.SubString( $basePathLength )
            if( $lastChangedInAzure = $existingItems[ $itemPath ] )
            {
                if( $lastChangedInAzure -gt $file.LastWriteTimeUTC )
                {
                    if( -Not $force)
                    {
                        Write-Warning -Message "Not syncing `"$itemPath`" because Azure copy is newer ($(Get-Date -Date $lastChangedInAzure -Format G)) than local $(Get-Date -Date $file.LastWriteTimeUTC -Format G))"
                        $sync = $false
                    }
                    else
                    {
                        Write-Warning -Message "Forcing update of `"$itemPath`" even though Azure copy is newer ($(Get-Date -Date $lastChangedInAzure -Format G)) than local $(Get-Date -Date $file.LastWriteTimeUTC -Format G))"
                    }
                }
            }
            elseif( $existingOnly )
            {
                $sync = $false
                Write-Verbose -Message "Not syncing `"$itemPath`" as does not exist in Azure"
            }
            if( $sync )
            {
                Write-Verbose -Message "Initial sync of $($file.fullname) ..."
                $storageBlobContentParameters.File = $file.fullname
                ## need to include any sub folders so they get created in Azure so we strip off the path before the root of the folder we're monitoring
                $storageBlobContentParameters.Blob =  $itemPath
                $storageBlobContentParameters.Properties.ContentType = Get-MimeMapping -File $file.fullname

                if( -Not (  $storageContainer | Set-AzStorageBlobContent @storageBlobContentParameters ))
                {
                    Write-Warning -Message "$(Get-Date -Format G): error on intial syncing of $($storageBlobContentParameters.File)"
                }
            }
        }
    }
    if( -Not $nowatch )
    {
        try
        {
            $filesystemWatcher = New-Object -typename System.IO.FileSystemWatcher -argumentlist $folder, $filter -Property @{IncludeSubdirectories = $includeSubdirectories.IsPresent ; NotifyFilter = [IO.NotifyFilters]'Size, FileName, DirectoryName, CreationTime, LastWrite'} -ErrorAction Stop

            Unregister-Event -SourceIdentifier FileCreated -Force -ErrorAction SilentlyContinue
            Unregister-Event -SourceIdentifier FileChanged -Force -ErrorAction SilentlyContinue
            Unregister-Event -SourceIdentifier FileRenamed -Force -ErrorAction SilentlyContinue
            ##Unregister-Event -SourceIdentifier FileDeleted -Force -ErrorAction SilentlyContinue

            Register-ObjectEvent -InputObject $filesystemWatcher -EventName Created -SourceIdentifier FileCreated
            Register-ObjectEvent -InputObject $filesystemWatcher -EventName Changed -SourceIdentifier FileChanged
            Register-ObjectEvent -InputObject $filesystemWatcher -EventName Renamed -SourceIdentifier FileRenamed
            ##Register-ObjectEvent -InputObject $filesystemWatcher -EventName Deleted -SourceIdentifier FileDeleted

            Write-Verbose "$(Get-Date -Format G) Waiting on folder $folder\$filter for $minutesToRun minutes"

            [int]$timeout = -1
            $endTime = $null

            if( $minutesToRun -gt 0 )
            {
                $endTime = [datetime]::Now.AddSeconds( $minutesToRun * 60 )
                Write-Verbose -Message "$(Get-Date -Format G): end time is $(Get-Date -Date $endTime -Format G)"
            }

            $lastEvent = $null

            do
            {
                if( $endTime )
                {
                    if( ( $timeout = ($endTime - [datetime]::Now).TotalSeconds ) -le 0 )
                    {
                        break
                    }
                    Write-Verbose -Message "Wait for $timeout seconds"
                }
                if( $eventRaised = Wait-Event -Timeout $timeout )
                {
                    ## We waited for any event so ensure it is ours and not an old one
                    if( $eventRaised.Sender.ToString() -eq $filesystemWatcher.GetType().fullname -And $eventRaised.SourceIdentifier -in @( 'FileCreated' , 'FileChanged' ) -And $eventRaised.TimeGenerated -ge $startTime )
                    {
                        ## have seen multiple events arrive for a single file system change so check they aren't
                        if( -not $lastEvent -or ($eventRaised.TimeGenerated - $lastEvent.TimeGenerated).TotalMilliseconds -ge $millisecondsApart )
                        {
                            [bool]$sync = $true
                            ## need to include any sub folders so they get created in Azure so we strip off the path before the root of the folder we're monitoring
                            [string]$itemPath = $eventRaised.SourceEventArgs.FullPath.SubString( $basePathLength )
                            if( -Not $existingOnly -or $existingItems[ $itemPath ] )
                            {
                                $changeType = $eventRaised.SourceEventArgs.ChangeType
                                $timeStamp = $eventRaised.TimeGenerated
                                $localFileProperties = Get-ItemProperty -Path $eventRaised.SourceEventArgs.FullPath
                                if( ( $localFileProperties.Attributes -band [System.IO.fileattributes]::Directory) -ne [System.IO.fileattributes]::Directory )
                                {
                                    Write-Verbose "The file '$($eventRaised.SourceEventArgs.FullPath)' size $([math]::Round( $localFileProperties.Length / 1KB, 1)) KB was $changeType at $(Get-Date -Date $timeStamp -Format G)"

                                    $storageBlobContentParameters.File = $eventRaised.SourceEventArgs.FullPath
                                    $storageBlobContentParameters.Blob =  $itemPath
                                    $storageBlobContentParameters.Properties.ContentType = Get-MimeMapping -File $eventRaised.SourceEventArgs.FullPath

                                    if( -Not ( $storageContainer | Set-AzStorageBlobContent @storageBlobContentParameters ))
                                    {
                                        Write-Warning -Message "$(Get-Date -Format G): error syncing $($storageBlobContentParameters.File)"
                                    }
                                }
                                else ## it's a directory which we won't sync directly as we'll wait for files as folders don't technically exist
                                {
                                    Write-Verbose "The folder '$($eventRaised.SourceEventArgs.FullPath)' was $changeType at $(Get-Date -Date $timeStamp -Format G)"
                                }
                            }
                            else
                            {
                                Write-Verbose -Message "Not syncing `"$itemPath`" as does not exist in Azure"
                            }
                        }
                        else
                        {
                            Write-Verbose -Message "Ignoring"
                        }
                        $lastEvent = $eventRaised
                    }
                    else
                    {
                        Write-Verbose "Got no relevant event - $($eventRaised.SourceIdentifier) @ $(Get-Date -Date $eventRaised.TimeGenerated -Format G)"
                    }
                    $eventRaised | Remove-Event
                }
            } while( -Not $endTime -or [datetime]::Now -lt $endTime )
        }
        catch
        {
            throw $_
        }
        finally
        {
            Unregister-Event -SourceIdentifier FileCreated -Force -ErrorAction SilentlyContinue
            Unregister-Event -SourceIdentifier FileChanged -Force -ErrorAction SilentlyContinue
            Unregister-Event -SourceIdentifier FileRenamed -Force -ErrorAction SilentlyContinue
            ##Unregister-Event -SourceIdentifier FileDeleted -Force -ErrorAction SilentlyContinue
        }
    }
}

# SIG # Begin signature block
# MIIZsAYJKoZIhvcNAQcCoIIZoTCCGZ0CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUUTeTsiNLxaRbeyMY8V0WFZ8P
# wh6gghS+MIIE/jCCA+agAwIBAgIQDUJK4L46iP9gQCHOFADw3TANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgVGltZXN0YW1waW5nIENBMB4XDTIxMDEwMTAwMDAwMFoXDTMxMDEw
# NjAwMDAwMFowSDELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMu
# MSAwHgYDVQQDExdEaWdpQ2VydCBUaW1lc3RhbXAgMjAyMTCCASIwDQYJKoZIhvcN
# AQEBBQADggEPADCCAQoCggEBAMLmYYRnxYr1DQikRcpja1HXOhFCvQp1dU2UtAxQ
# tSYQ/h3Ib5FrDJbnGlxI70Tlv5thzRWRYlq4/2cLnGP9NmqB+in43Stwhd4CGPN4
# bbx9+cdtCT2+anaH6Yq9+IRdHnbJ5MZ2djpT0dHTWjaPxqPhLxs6t2HWc+xObTOK
# fF1FLUuxUOZBOjdWhtyTI433UCXoZObd048vV7WHIOsOjizVI9r0TXhG4wODMSlK
# XAwxikqMiMX3MFr5FK8VX2xDSQn9JiNT9o1j6BqrW7EdMMKbaYK02/xWVLwfoYer
# vnpbCiAvSwnJlaeNsvrWY4tOpXIc7p96AXP4Gdb+DUmEvQECAwEAAaOCAbgwggG0
# MA4GA1UdDwEB/wQEAwIHgDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB/wQMMAoGCCsG
# AQUFBwMIMEEGA1UdIAQ6MDgwNgYJYIZIAYb9bAcBMCkwJwYIKwYBBQUHAgEWG2h0
# dHA6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzAfBgNVHSMEGDAWgBT0tuEgHf4prtLk
# YaWyoiWyyBc1bjAdBgNVHQ4EFgQUNkSGjqS6sGa+vCgtHUQ23eNqerwwcQYDVR0f
# BGowaDAyoDCgLoYsaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJl
# ZC10cy5jcmwwMqAwoC6GLGh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9zaGEyLWFz
# c3VyZWQtdHMuY3JsMIGFBggrBgEFBQcBAQR5MHcwJAYIKwYBBQUHMAGGGGh0dHA6
# Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBPBggrBgEFBQcwAoZDaHR0cDovL2NhY2VydHMu
# ZGlnaWNlcnQuY29tL0RpZ2lDZXJ0U0hBMkFzc3VyZWRJRFRpbWVzdGFtcGluZ0NB
# LmNydDANBgkqhkiG9w0BAQsFAAOCAQEASBzctemaI7znGucgDo5nRv1CclF0CiNH
# o6uS0iXEcFm+FKDlJ4GlTRQVGQd58NEEw4bZO73+RAJmTe1ppA/2uHDPYuj1UUp4
# eTZ6J7fz51Kfk6ftQ55757TdQSKJ+4eiRgNO/PT+t2R3Y18jUmmDgvoaU+2QzI2h
# F3MN9PNlOXBL85zWenvaDLw9MtAby/Vh/HUIAHa8gQ74wOFcz8QRcucbZEnYIpp1
# FUL1LTI4gdr0YKK6tFL7XOBhJCVPst/JKahzQ1HavWPWH1ub9y4bTxMd90oNcX6X
# t/Q/hOvB46NJofrOp79Wz7pZdmGJX36ntI5nePk2mOHLKNpbh6aKLzCCBTAwggQY
# oAMCAQICEAQJGBtf1btmdVNDtW+VUAgwDQYJKoZIhvcNAQELBQAwZTELMAkGA1UE
# BhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2lj
# ZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENBMB4X
# DTEzMTAyMjEyMDAwMFoXDTI4MTAyMjEyMDAwMFowcjELMAkGA1UEBhMCVVMxFTAT
# BgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEx
# MC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVkIElEIENvZGUgU2lnbmluZyBD
# QTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAPjTsxx/DhGvZ3cH0wsx
# SRnP0PtFmbE620T1f+Wondsy13Hqdp0FLreP+pJDwKX5idQ3Gde2qvCchqXYJawO
# eSg6funRZ9PG+yknx9N7I5TkkSOWkHeC+aGEI2YSVDNQdLEoJrskacLCUvIUZ4qJ
# RdQtoaPpiCwgla4cSocI3wz14k1gGL6qxLKucDFmM3E+rHCiq85/6XzLkqHlOzEc
# z+ryCuRXu0q16XTmK/5sy350OTYNkO/ktU6kqepqCquE86xnTrXE94zRICUj6whk
# PlKWwfIPEvTFjg/BougsUfdzvL2FsWKDc0GCB+Q4i2pzINAPZHM8np+mM6n9Gd8l
# k9ECAwEAAaOCAc0wggHJMBIGA1UdEwEB/wQIMAYBAf8CAQAwDgYDVR0PAQH/BAQD
# AgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMDMHkGCCsGAQUFBwEBBG0wazAkBggrBgEF
# BQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEMGCCsGAQUFBzAChjdodHRw
# Oi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0Eu
# Y3J0MIGBBgNVHR8EejB4MDqgOKA2hjRodHRwOi8vY3JsNC5kaWdpY2VydC5jb20v
# RGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMDqgOKA2hjRodHRwOi8vY3JsMy5k
# aWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsME8GA1UdIARI
# MEYwOAYKYIZIAYb9bAACBDAqMCgGCCsGAQUFBwIBFhxodHRwczovL3d3dy5kaWdp
# Y2VydC5jb20vQ1BTMAoGCGCGSAGG/WwDMB0GA1UdDgQWBBRaxLl7KgqjpepxA8Bg
# +S32ZXUOWDAfBgNVHSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzANBgkqhkiG
# 9w0BAQsFAAOCAQEAPuwNWiSz8yLRFcgsfCUpdqgdXRwtOhrE7zBh134LYP3DPQ/E
# r4v97yrfIFU3sOH20ZJ1D1G0bqWOWuJeJIFOEKTuP3GOYw4TS63XX0R58zYUBor3
# nEZOXP+QsRsHDpEV+7qvtVHCjSSuJMbHJyqhKSgaOnEoAjwukaPAJRHinBRHoXpo
# aK+bp1wgXNlxsQyPu6j4xRJon89Ay0BEpRPw5mQMJQhCMrI2iiQC/i9yfhzXSUWW
# 6Fkd6fp0ZGuy62ZD2rOwjNXpDd32ASDOmTFjPQgaGLOBm0/GkxAG/AeB+ova+YJJ
# 92JuoVP6EpQYhS6SkepobEQysmah5xikmmRR7zCCBTEwggQZoAMCAQICEAqhJdbW
# Mht+QeQF2jaXwhUwDQYJKoZIhvcNAQELBQAwZTELMAkGA1UEBhMCVVMxFTATBgNV
# BAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEkMCIG
# A1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENBMB4XDTE2MDEwNzEyMDAw
# MFoXDTMxMDEwNzEyMDAwMFowcjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lD
# ZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTExMC8GA1UEAxMoRGln
# aUNlcnQgU0hBMiBBc3N1cmVkIElEIFRpbWVzdGFtcGluZyBDQTCCASIwDQYJKoZI
# hvcNAQEBBQADggEPADCCAQoCggEBAL3QMu5LzY9/3am6gpnFOVQoV7YjSsQOB0Uz
# URB90Pl9TWh+57ag9I2ziOSXv2MhkJi/E7xX08PhfgjWahQAOPcuHjvuzKb2Mln+
# X2U/4Jvr40ZHBhpVfgsnfsCi9aDg3iI/Dv9+lfvzo7oiPhisEeTwmQNtO4V8CdPu
# XciaC1TjqAlxa+DPIhAPdc9xck4Krd9AOly3UeGheRTGTSQjMF287DxgaqwvB8z9
# 8OpH2YhQXv1mblZhJymJhFHmgudGUP2UKiyn5HU+upgPhH+fMRTWrdXyZMt7HgXQ
# hBlyF/EXBu89zdZN7wZC/aJTKk+FHcQdPK/P2qwQ9d2srOlW/5MCAwEAAaOCAc4w
# ggHKMB0GA1UdDgQWBBT0tuEgHf4prtLkYaWyoiWyyBc1bjAfBgNVHSMEGDAWgBRF
# 66Kv9JLLgjEtUYunpyGd823IDzASBgNVHRMBAf8ECDAGAQH/AgEAMA4GA1UdDwEB
# /wQEAwIBhjATBgNVHSUEDDAKBggrBgEFBQcDCDB5BggrBgEFBQcBAQRtMGswJAYI
# KwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3
# aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9v
# dENBLmNydDCBgQYDVR0fBHoweDA6oDigNoY0aHR0cDovL2NybDQuZGlnaWNlcnQu
# Y29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDA6oDigNoY0aHR0cDovL2Ny
# bDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDBQBgNV
# HSAESTBHMDgGCmCGSAGG/WwAAgQwKjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cu
# ZGlnaWNlcnQuY29tL0NQUzALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggEB
# AHGVEulRh1Zpze/d2nyqY3qzeM8GN0CE70uEv8rPAwL9xafDDiBCLK938ysfDCFa
# KrcFNB1qrpn4J6JmvwmqYN92pDqTD/iy0dh8GWLoXoIlHsS6HHssIeLWWywUNUME
# aLLbdQLgcseY1jxk5R9IEBhfiThhTWJGJIdjjJFSLK8pieV4H9YLFKWA1xJHcLN1
# 1ZOFk362kmf7U2GJqPVrlsD0WGkNfMgBsbkodbeZY4UijGHKeZR+WfyMD+NvtQEm
# tmyl7odRIeRYYJu6DC0rbaLEfrvEJStHAgh8Sa4TtuF8QkIoxhhWz0E0tmZdtnR7
# 9VYzIi8iNrJLokqV2PWmjlIwggVPMIIEN6ADAgECAhAE/eOq2921q55B9NnVIXVO
# MA0GCSqGSIb3DQEBCwUAMHIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2Vy
# dCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lD
# ZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EwHhcNMjAwNzIwMDAw
# MDAwWhcNMjMwNzI1MTIwMDAwWjCBizELMAkGA1UEBhMCR0IxEjAQBgNVBAcTCVdh
# a2VmaWVsZDEmMCQGA1UEChMdU2VjdXJlIFBsYXRmb3JtIFNvbHV0aW9ucyBMdGQx
# GDAWBgNVBAsTD1NjcmlwdGluZ0hlYXZlbjEmMCQGA1UEAxMdU2VjdXJlIFBsYXRm
# b3JtIFNvbHV0aW9ucyBMdGQwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQCvbSdd1oAAu9rTtdnKSlGWKPF8g+RNRAUDFCBdNbYbklzVhB8hiMh48LqhoP7d
# lzZY3YmuxztuPlB7k2PhAccd/eOikvKDyNeXsSa3WaXLNSu3KChDVekEFee/vR29
# mJuujp1eYrz8zfvDmkQCP/r34Bgzsg4XPYKtMitCO/CMQtI6Rnaj7P6Kp9rH1nVO
# /zb7KD2IMedTFlaFqIReT0EVG/1ZizOpNdBMSG/x+ZQjZplfjyyjiYmE0a7tWnVM
# Z4KKTUb3n1CTuwWHfK9G6CNjQghcFe4D4tFPTTKOSAx7xegN1oGgifnLdmtDtsJU
# OOhOtyf9Kp8e+EQQyPVrV/TNAgMBAAGjggHFMIIBwTAfBgNVHSMEGDAWgBRaxLl7
# KgqjpepxA8Bg+S32ZXUOWDAdBgNVHQ4EFgQUTXqi+WoiTm5fYlDLqiDQ4I+uyckw
# DgYDVR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMHcGA1UdHwRwMG4w
# NaAzoDGGL2h0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9zaGEyLWFzc3VyZWQtY3Mt
# ZzEuY3JsMDWgM6Axhi9odHRwOi8vY3JsNC5kaWdpY2VydC5jb20vc2hhMi1hc3N1
# cmVkLWNzLWcxLmNybDBMBgNVHSAERTBDMDcGCWCGSAGG/WwDATAqMCgGCCsGAQUF
# BwIBFhxodHRwczovL3d3dy5kaWdpY2VydC5jb20vQ1BTMAgGBmeBDAEEATCBhAYI
# KwYBBQUHAQEEeDB2MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5j
# b20wTgYIKwYBBQUHMAKGQmh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdp
# Q2VydFNIQTJBc3N1cmVkSURDb2RlU2lnbmluZ0NBLmNydDAMBgNVHRMBAf8EAjAA
# MA0GCSqGSIb3DQEBCwUAA4IBAQBT3M71SlOQ8vwM2txshp/XDvfoKBYHkpFCyanW
# aFdsYQJQIKk4LOVgUJJ6LAf0xPSN7dZpjFaoilQy8Ajyd0U9UOnlEX4gk2J+z5i4
# sFxK/W2KU1j6R9rY5LbScWtsV+X1BtHihpzPywGGE5eth5Q5TixMdI9CN3eWnKGF
# kY13cI69zZyyTnkkb+HaFHZ8r6binvOyzMr69+oRf0Bv/uBgyBKjrmGEUxJZy+00
# 7fbmYDEclgnWT1cRROarzbxmZ8R7Iyor0WU3nKRgkxan+8rzDhzpZdtgIFdYvjeO
# c/IpPi2mI6NY4jqDXwkx1TEIbjUdrCmEfjhAfMTU094L7VSNMYIEXDCCBFgCAQEw
# gYYwcjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UE
# CxMQd3d3LmRpZ2ljZXJ0LmNvbTExMC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1
# cmVkIElEIENvZGUgU2lnbmluZyBDQQIQBP3jqtvdtaueQfTZ1SF1TjAJBgUrDgMC
# GgUAoHgwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYK
# KwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG
# 9w0BCQQxFgQU3jdSB46m2jceHCcxG9uUCALfWUIwDQYJKoZIhvcNAQEBBQAEggEA
# doycQFJ3kIQTMkKKqm8ZZVhKQQp/oyGbXl5qqfJN5KCWGLn2+wJFCRr8SKWAAZn1
# d36HJ4VpqXD5x57orQwqcT4bWl/b8bWEssHYkIyDtDxSXENMmWmzGkD24OZPIJxt
# uZ9EbHGoukqTHuZzNasccJON1YkaDxjAcfhPf2isSAP4e5WbHmJJ8CASyGf8rMui
# MEI3uOQwsUvfmyF7+7ZHwPhPfbcF3bEuk3yJjae5TP/VA0wlSqXzDkt7wFtBEuiY
# SzdTUFWHS3iKgAepgz0ucAa3GqhhGU7DzfvYjHl56dQqQF3pCkqhg7Yo3gLSxErt
# y2j9kS9VFpd2nj02Jm0K6qGCAjAwggIsBgkqhkiG9w0BCQYxggIdMIICGQIBATCB
# hjByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQL
# ExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3Vy
# ZWQgSUQgVGltZXN0YW1waW5nIENBAhANQkrgvjqI/2BAIc4UAPDdMA0GCWCGSAFl
# AwQCAQUAoGkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUx
# DxcNMjIwMzE0MDk1MzQyWjAvBgkqhkiG9w0BCQQxIgQgggIx1D5+VgwRh5BkBKqZ
# tPvT7ZSnOS+/Fkydfy3Djp8wDQYJKoZIhvcNAQEBBQAEggEAUAGMfj1FCs4TpNxQ
# D0HExX5/2uF+WhZP+gbmzCNuHGQtpGZH6SF9toOAWNgHG67LX/LHHOC+jGKrLUQK
# IcjG1pW1FBNLNhRaJ04PJWxv9tKxXYpPfxDusRbTiUH5AnHCkuEqsks7mfIm/0Pb
# RPA2KSvgRK/8km7KGXvBd1qUS0XSS/AGVPw9n8DsSs+8lUqVUc6ZqgKsJDMmDRxG
# 32s7PFhM7CUcyR8ItQSYam1730nvh1W3ZmWzZla4DeVM0H6DdW9PLsezMgo6mowv
# Etpjc5oh2E/M6E9HhtUkggXk63V0UyX9rHXF4vUPfLQXVUtHTsN22FAi2hIQqh8p
# ET1uvQ==
# SIG # End signature block
