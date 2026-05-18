<#
.SYNOPSIS
	Show session detail
.NOTES
	Modification History:
	2026/02/13  Guy Leech  Script born
	2026/02/25  Guy Leech  Metadata, IP addresses and FSlogix disk information added
    2026/03/06  Guy Leech  Added Read-Host if run from explorer
    2026/03/09  Guy Leech  Added internal IPv6 address output
    2026/03/12  Guy Leech  Changes to detect if run from explorer. Added FSlogix profile size. Filter out IPv4 with no address
    2026/04/10  Guy Leech  Added AD lookup, set width of console if run from explorer
    2026/04/15  Guy Leech  Added dsregcmd /status. Added AZ metadata
    2026/05/05  Guy Leech  Added count of stored credentials
    2026/05/17  Guy Leech  Added disk space
    2026/05/18  Guy Leech  Enhanced AZ metadata output
#>

[string]$runFromExplorerRegex = '-ExecutionPolicy\b|-File\b'
[bool]$runFromExplorer = $false

if( $host.Name -ieq 'ConsoleHost' )
{
    <# from explorer:
    Invocation: if((Get-ExecutionPolicy ) -ne 'AllSigned') { Set-ExecutionPolicy -Scope Process Bypass }; & 'C:\Users\blah\Guys Scripts\Session Information.ps1'
    Command line: "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" "-Command" "if((Get-ExecutionPolicy ) -ne 'AllSigned') { Set-ExecutionPolicy -Scope Process Bypass }; & 'C:\Users\blah\Guys Scripts\Session Information.ps1'"
    #>

    $thisProcess = $null
    $thisProcess = Get-CimInstance -ClassName win32_process -Filter "ProcessId = $pid" -ErrorAction SilentlyContinue

    if( $null -ne $thisProcess -and ( $thisProcess.CommandLine -match $runFromExplorerRegex -or ( $myinvocation.ScriptLineNumber -eq 0 -and $myinvocation.OffsetInLine -eq 0 ))) ## these will be at least 1 when interactive
    {
        $parentProcess = $null
        $parentProcess = Get-Process -Id  $thisProcess.ParentProcessId -ErrorAction SilentlyContinue
        if( $null -ne $parentProcess -and $parentProcess.Name -ieq 'explorer' ) ## cannot use this alone since an existing PS prompt's parent may also be explorer
        {
            $runFromExplorer = $true
            ## make display wider
            [int]$outputWidth = 150
            try
            {
                if( ( $PSWindow = (Get-Host).UI.RawUI ) -and ( $WideDimensions = $PSWindow.BufferSize ) )
                {
                    $WideDimensions.Width = $outputWidth
                    $PSWindow.BufferSize = $WideDimensions
                    $windowSize = $PSWindow.WindowSize
                    $windowSize.Width = $outputWidth
                    $PSWindow.WindowSize = $windowSize
                }
            }
            catch
            {
                ## not much we can do but will hide the error since it is not fundamental to script functionality, just output
                Write-Warning -Message "Failed to set output width to $($WideDimensions.width) : $_"
            }
        }
    }
}

$oldProgressPreference = $ProgressPreference
$ProgressPreference = 'SilentlyContinue'
$info = Get-ComputerInfo ## -Property OsTotalVisibleMemorySize,CsNumberOfLogicalProcessors
$graphics = Get-CimInstance Win32_VideoController | Select-Object Name, AdapterCompatibility, VideoProcessor, @{Name='DedicatedVRAM_GB';Expression={"{0:N2}" -f ($_.AdapterRAM / 1GB)}}
$ProgressPreference = $oldProgressPreference

"$(Get-Date -Format 'G'): Machine $($env:COMPUTERNAME), district $($env:COMPUTERNAME -replace '^k(\w{6}).*$' , '$1'), booted $((gcim win32_operatingsystem).LastBootupTime.ToString('G')) : pid $pid`n"

"CPUs = {0}, memory = {1:N2} GB" -f $info.CsNumberOfLogicalProcessors , ($info.OsTotalVisibleMemorySize / 1MB)
''
'Video Controller:'
($graphics|Out-String).TrimEnd()

(Get-Volume | Where-Object Size -gt 1GB | Select-Object -Property DriveLetter,FileSystemLabel,@{n='Size GB';e={[int]($_.Size / 1GB)}},@{n='Remaining GB';e={[int]($_.SizeRemaining / 1GB)}}|ft -auto|Out-String).TrimEnd()
''
"Profile Disk: $(((Get-disk).location | where { $_ -notmatch '^(PCI Slot|Integrated)' } | Select @{n='Path';e={$_}},@{n='Size GB';e={$script:pd=dir $_ -EA 0;'{0:N2}' -f ($script:pd.Length / 1GB)}},@{n='Created';e={$script:pd.CreationTime}},@{n='Modified';e={$script:pd.LastWriteTime}}|ft -auto|Out-String).TrimEnd())"
"Profile size limit: $((Get-ItemPropertyValue -Path "hklm:\SOFTWARE\FSLogix\Profiles" -Name SizeInMBs)/1024) GB"
"Number of stored credentials : $( (cmdkey /list|sls \bTarget:|measure).Count)"
''
"External IP address   : $((iwr -usebasicparsing http://icanhazip.com) -replace "`r?`n")"
"Internal IP address   : $((gip).IPv4Address.IPAddress|where { $_ -notmatch '^169\.254\.' } )"
"Internal IPv6 address : $(Get-NetIPAddress -AddressFamily IPv6 -EA 0 |Where IPAddress -ne '::1' -EA 0|select -expand IPAddress -EA 0)"
''
$searcher = [adsisearcher]"(&(objectClass=computer)(name=$env:COMPUTERNAME))"
$result = $null
try
{
    $result = $searcher.FindOne()
    "Distinguised name  : $($result.Properties['distinguishedname'][0])"
    "Computer in groups :"
    $result.Properties['memberof']
}
catch
{
    Write-Warning "Failed to find computer $env:COMPUTERNAME in AD: $_"
}
"Domain             : $(([adsi]"LDAP://RootDSE").defaultNamingContext) , logon server $Env:LOGONSERVER"

"dsregcmd /status"
dsregcmd /status|Select-String 'Joined : | Tenant.*: |Domain|Executing Account Name : '
''
quser.exe
''
if( Test-Path -Path "$env:ProgramFiles\itopia\Scripts\Labs-BuildFunctions.psm1" -PathType Leaf )
{
	Import-Module -Name "$env:ProgramFiles\itopia\Scripts\Labs-BuildFunctions.psm1" -InformationAction silentlycontinue
	(Get-SessionMetadata).getenumerator()|select Name,Value|sort Name|format-table -Auto
}
else ## see if Azure
{
    [hashtable]$results = @{}
    [string]$ImdsServer = '169.254.169.254'
    [string]$apiversion = '2017-03-01'
    [hashtable]$headers =  @{ "Metadata" = "true" }
    
    ## https://stackoverflow.com/posts/66345592/revisions
    $WebSession = New-Object -TypeName Microsoft.PowerShell.Commands.WebRequestSession
    $WebSession.Proxy = New-object -Typename System.Net.WebProxy

    try
    {
        [array]$latestVersions = @( Invoke-RestMethod -Headers $headers -URI "http://$ImdsServer/metadata/versions?api-version=$apiversion" -WebSession $WebSession | Select-Object -ExpandProperty apiversions -Last 1 )
        if( $latestVersions -and $latestVersions.Count -gt 0 )
        {
            $apiversion = $latestVersions[ -1 ]
        }
    }
    catch
    {
        Write-Warning -Message "Failed to get later version of api than $apiversion - $_"
    }
    Write-Verbose -Message "Using api version $apiversion"
    $metadata = $null
    ## https://learn.microsoft.com/en-us/azure/virtual-machines/windows/instance-metadata-service?tabs=windows
    [string]$uri = "http://$ImdsServer/metadata/instance/compute?api-version=$apiversion"
    try
    {
        $metadata = Invoke-RestMethod -Headers $headers -URI $uri -WebSession $WebSession
        'Azure metadata:'
            ($metadata|Select-Object -Property Resource*,SubscriptionId,vmSize|Format-List -Property *|Out-String).TrimEnd()
        ''
        'Azure image:'
            ($metadata.storageProfile.imageReference.psobject.Properties|Where-Object { -not [string]::IsNullOrEmpty( $_.value ) }|sort name|select name,value|Out-String).TrimEnd()
            ($metadata.storageprofile.osdisk|Select-Object -Property manageddisk|Format-List -Property *|Out-String).TrimEnd()
        ''
        'Azure tags:'
            ($metadata.tagsList|select Name,Value|sort name|Format-Table -AutoSize|Out-String).TrimEnd()
    }
    catch
    {
        Write-Warning -Message "Failed to get Azure metadata from $uri - $_"
    }
}

## if parent is explorer then prompt to continue as may have been run via right click and conhost will disappear on exit losing output
if( $runFromExplorer )
{
    $null = Read-Host -Prompt "Hit Enter to exit "
}
