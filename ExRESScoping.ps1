<#
 ==========[DISCLAIMER]===========================================================================================================
  This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment.  
  THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, 
  INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  
  We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object
  code form of the Sample Code, provided that You agree: (i) to not use Our name, logo, or trademarks to market Your software 
  product in which the Sample Code is embedded; (ii) to include a valid copyright notice on Your software product in which the 
  Sample Code is embedded; and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims or 
  lawsuits, including attorneys’ fees, that arise or result from the use or distribution of the Sample Code.
 =================================================================================================================================
 
 Authors: Dmitriy Plokhih and Dmitry Goncharov
 Updated Script by Dmitry Goncharov at https://github.com/ExRES/ExRESScoping
 Version 2.2-2018.12.06
#>

#HTML header for ExRESScoping.html report 
$Header = @"
<style>
TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #6495ED;}
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
.odd  { background-color:#ffffff; }
.even { background-color:#dddddd; }
</style>
"@

function Create-ZipFile
{
<#
.Synopsis
   Archive a report to a ZIP file
.DESCRIPTION
   Archive ExRESScoping report (xml and html files) to a ZIP file
.EXAMPLE
   Archive reports in 'Documents' folder and create ExRESScoping.zip archive
   ---------------------------------
   $MyDocsPath = [Environment]::GetFolderPath("MyDocuments")
   Create-ZipFile -dir $MyDocsPath
.PARAMETER dir
   Directory of collected xml files
.PARAMETER mask
   Wildcard mask for getting source files files for a archive. The default value covers ExRESScoping.html and ExRESScoping.xml files.
.PARAMETER zipFileName
   Name of a archive file. The default value is ExRESScoping.zip

#>
param
( [string]$dir, #Directory of collected xml files
  [string]$mask="ExRESScoping*.*l", #wildcard mask for getting ExRESScoping.html and ExRESScoping.xml files 
  [string]$zipFileName="ExRESScoping.zip" # name of ZIP archive

)
#Full path to a ZIP file
$zipFile = Join-Path -Path $dir -ChildPath $zipFilename
#Wildcard mask for filtering source files for an archive
$searchStr = Join-Path -Path $dir -ChildPath $mask

#Prepare zip file. An old file will be overwritten
    set-content $zipFile ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18))
#Remove "ReadOnly" flag for a ZIP file
    (dir $zipFile).IsReadOnly = $false  

#Get object for ZIP file management
$shellApplication = new-object -com shell.application
#Object for managing the prepared ZIP file
$zipPackage = $shellApplication.NameSpace($zipFile)

#Get files by using provided mask
$files = Get-ChildItem -Path $searchStr | where{! $_.PSIsContainer}

foreach($file in $files) { 
    #Copy every file in "$Files" array to a archive
    $zipPackage.CopyHere($file.FullName)
#using this method, sometimes files can be 'skipped'
#this 'while' loop checks each file is added before moving to the next
    while($zipPackage.Items().Item($file.name) -eq $null){
        Start-sleep -seconds 1
    }
}
}


function Invoke-RunspaceJob
{
<#
.Synopsis
   Creation of parallel background processes
.DESCRIPTION
   Creation of parallel background processes and execution of a provided scriptblock in a separate runspace.
.PARAMETER InputObject 
   One or more process objects
.PARAMETER ScriptBlock 
   PowerShell code for execution
.PARAMETER ThrottleLimit
   Number of asynchronously running tasks
.PARAMETER Timeout
   Number of asynchronously running tasks
.PARAMETER ShowProgress
   Display progress bar
#>
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   Position=1)]
        [ValidateNotNullOrEmpty()]
        [PSObject[]]
        $InputObject, #process objects

        [Parameter(Mandatory=$true, 
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.ScriptBlock]
        $ScriptBlock, #PowerShell code for execution

        [Parameter(Position=2)]
        [Int32]
        $ThrottleLimit = 32, #Number of asynchronously running tasks

        [Parameter(Position=3)]
        [Int32]
        $Timeout, #Timeout

        [Parameter(Position=5)]
        [switch]
        $ShowProgress, #Display progress bar

        [Parameter(Position=4)]
        [ValidateScript({$_ | ForEach-Object -Process {Get-Variable -Name $_}})]
        [string[]]
        $SharedVariables #Collection of shared variables for importing into a runspace
    )

    Begin
    {
        #region Creating initial variables
        #Preparation of 'runspacetimers' hashtable for multiple threads. It's used for a runspace identification
        $runspacetimers = [HashTable]::Synchronized(@{}) 
        $SharedVariables += 'runspacetimers' #Add 'runspacetimers' variable to a list of shared variables 
        $runspaces = New-Object -TypeName System.Collections.ArrayList #Array for runspaces
        $bgRunspaceCounter = 0 #Initial value for runspace counter
        #endregion Creating initial variables

        #region Creating initial session state and runspace pool
        Write-Verbose -Message "Creating initial session state"
        #Creation of SessionState object
        $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
        #Add every variable in SharedVariables collection into SessionState object
        foreach ($ExternalVariable in $SharedVariables)
        {
            Write-Verbose -Message ('Adding variable ${0} to initial session state' -f $ExternalVariable)
            $iss.Variables.Add((New-Object -TypeName System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $ExternalVariable, (Get-Variable -Name $ExternalVariable -ValueOnly), ''))
        }
        Write-Verbose "Creating runspace pool with Throttle Limit $ThrottleLimit"
        #Creation of runspace pool by using previously created 'SessionState' object and 'Throttle limit'
        $rp = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $ThrottleLimit, $iss, $Host)
        $rp.Open()
        #endregion Creating initial session state and runspace pool

        #region Append timeout tracking code at the begining of scriptblock
        #The following scriptblock will be added at the beginning of provided scriptblock 
        $ScriptStart = {
            [CmdletBinding()]
            Param
            (
                [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   Position=0)]
                $_, # Object from pipeline

                [Parameter(Position=1)]
                [ValidateNotNullOrEmpty()]
                [int]
                $bgRunspaceID #Runspace ID
            )
            $runspacetimers.$bgRunspaceID = Get-Date #Assignment of current date and time to Runspace ID key
        }

        #Combining of 'ScriptStart' scriptblock and provided scriptblock in 'ScriptBlock' variable
        $ScriptBlock = [System.Management.Automation.ScriptBlock]::Create($ScriptStart.ToString() + $ScriptBlock.ToString())
        #endregion Append timeout tracking code at the begining of scriptblock

        #region Runspace status tracking and result retrieval function
        #Internal function for waiting for completion of runspaces
        function Get-Result
        {
            [CmdletBinding()]
            Param
            (
                [switch]$Wait #To wait for completion of a runspace
            )
            do
            {
                $More = $false
                #Check runspace state for every runspace 
                foreach ($runspace in $runspaces)
                {
                    #Get start time of a runspace
                    $StartTime = $runspacetimers.($runspace.ID)

                    #Release resources for all completed runspaces
                    if ($runspace.Handle.isCompleted)
                    {
                        Write-Verbose -Message ('Thread done for {0}' -f $runspace.IObject)
                        $runspace.PowerShell.EndInvoke($runspace.Handle)
                        $runspace.PowerShell.Dispose()
                        $runspace.PowerShell = $null
                        $runspace.Handle = $null
                    }
                    elseif ($runspace.Handle -ne $null)
                    {
                        #Set flag for waiting for completion of runspaces
                        $More = $true
                    }
                    #If Timeout is set then compare an elapsed time against Timeout                 
                    if ($Timeout -and $StartTime)
                    {
                    #If an elapsed time exceeded Timeout then release resources for a runspace
                        if ((New-TimeSpan -Start $StartTime).TotalMinutes -ge $Timeout)
                        {
                            Write-Warning -Message ('Timeout {0}' -f $runspace.IObject)
                            $runspace.PowerShell.Dispose()
                            $runspace.PowerShell = $null
                            $runspace.Handle = $null
                        }
                    }
                }
                #if More flag is set and "Wait" parameter is provided then pause for 100ms
                if ($More -and $PSBoundParameters['Wait'])
                {
                    Start-Sleep -Milliseconds 100
                }
                #Remove all runpaces without handles
                foreach ($threat in $runspaces.Clone())
                {
                    if ( -not $threat.handle)
                    {
                        Write-Verbose -Message ('Removing {0}' -f $threat.IObject)
                        $runspaces.Remove($threat)
                    }
                }
                #if ShowProgress parameter is provided then show a progress bar and display count of active runspaces 
                if ($ShowProgress)
                {
                #Parameters for Write-Progress cmdlet
                    $ProgressSplatting = @{
                        Activity = 'Working'
                        Status = 'Proccesing threads'
                        CurrentOperation = '{0} of {1} total threads done' -f ($bgRunspaceCounter - $runspaces.Count), $bgRunspaceCounter
                        PercentComplete = ($bgRunspaceCounter - $runspaces.Count) / $bgRunspaceCounter * 100
                    }
                #Display a progress bar
                    Write-Progress @ProgressSplatting
                }
            }
            while ($More -and $PSBoundParameters['Wait'])
        }
        #endregion Runspace status tracking and result retrieval function
    }
    Process
    {
        #Create a runspace for every provided object
        foreach ($Object in $InputObject)
        {
            #Increase current number of runspace ID
            $bgRunspaceCounter++
            #Create a runspace and assign ScriptBlock for execution, current runspace ID and provided object for processing
            $psCMD = [System.Management.Automation.PowerShell]::Create().AddScript($ScriptBlock).AddParameter('bgRunspaceID',$bgRunspaceCounter).AddArgument($Object)
            $psCMD.RunspacePool = $rp
            
            Write-Verbose -Message ('Starting {0}' -f $Object)
            #Add a created runspace to "Runspaces" array
            [void]$runspaces.Add(@{
                Handle = $psCMD.BeginInvoke()
                PowerShell = $psCMD
                IObject = $Object
                ID = $bgRunspaceCounter
           })
        #Check current state of runspaces
            Get-Result
        }
    }
    End
    {
        #Wait for completion for all runspaces
        Get-Result -Wait

        #if ShowProgress parameter is provided then show a progress bar
        if ($ShowProgress)
        {
            Write-Progress -Activity 'Working' -Status 'Done' -Completed
        }

        #Release resources for runspace pool
        Write-Verbose -Message "Closing runspace pool"
        $rp.Close()
        $rp.Dispose()
    }
}

function Get-DomainNetBIOSName
{
<#
.Synopsis
   Get NETBIOS name for a provided domain
.DESCRIPTION
   Get NETBIOS name for a provided domain
.PARAMETER Identity
   Domain's fqdn
#>
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [string[]]
        $Identity
    )

    Process
    {
        #Connect to AD and get NETBIOS name for every profided domain fqdn
        foreach ($Domain in $Identity)
        {
                #Preparation of connection to a AD configuration partition
                #Creation of DirectoryEntry object by using ADSI type accelerator
                $RootDSE = [ADSI]"LDAP://RootDSE"
                #Get path to Configuration partition for selected domain
                $ConfigNC = $RootDSE.Get("configurationNamingContext")
                #Root of AD search
                $ADSearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://CN=Partitions," + $ConfigNC)
                #AD search filter
                $SearchString="(&(objectclass=Crossref)(dnsRoot="+$Domain+")(netBIOSName=*))"
                #Create DirectorySearcher object by using the filter and the AD search root
                $Search = New-Object System.DirectoryServices.DirectorySearcher($ADSearchRoot,$SearchString)
                #Start search and get domain's NETBIOS name
                $NetBIOSName = ($Search.FindOne()).Properties["netbiosname"]
                #Send NETBIOS name to a function output
                Write-Output $NetBIOSName
        }
    }
}

function Get-ForestInfo
{
<#
.Synopsis
   Get info about current AD forest
.DESCRIPTION
   Get info about current AD forest: 
        - list of domains
        - AD sites
        - name of a root domain
        - Forest level
        - global catalogs 
.PARAMETER Identity
   Domain's fqdn
#>
    [CmdletBinding()]
    Param
    (
    )
    #Start a static method of provided class to get an info about the current forest
    [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
}


#Check if Exchange cmdlets are loaded in the current session
if (!(Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue))
{
    #If not loaded, check if Exchange Management Tools component is installed
	if (Test-Path "$($env:ExchangeInstallPath)bin\RemoteExchange.ps1")
	{
        #Prepare a connection to an Exchange server
		. "$($env:ExchangeInstallPath)bin\RemoteExchange.ps1"
        #Connect to an Exchange server and import Exchange cmdlets
		Connect-ExchangeServer -auto
	} else {
        #Otherwise, the script will be stopped
		throw "Exchange Management Shell cannot be loaded"
	}
}
#Prepare a variable for XML data
$xmlData = @()
#All objects in the forest are viewed and managed in this session
Set-ADServerSettings -ViewEntireForest:$true
#Get path to "Documents" folder
$MyDocsPath = [Environment]::GetFolderPath("MyDocuments")
#Compose full path for ExRESScoping.html in Documents folder 
$HTMLOutPath = Join-Path -Path $MyDocsPath -ChildPath "ExRESScoping.html"
#Compose full path for ExRESScoping.xml in Documents folder
$XMLOutPath = Join-Path -Path $MyDocsPath -ChildPath "ExRESScoping.xml"

try
{
    #Get AD Forest info
    $Forest = Get-ForestInfo -ErrorAction Stop
    #Convert AD Forest info to HTML code    
    $ForestHTML = $Forest | ConvertTo-Html -Property Name, ForestMode -Fragment -As List -PreContent "<h2>Forest:</h2>" | Out-String

}
catch
{
    #Stop the script in case of an error and display an error message
    throw "Failed to get AD Data: $($_.Exception.Message)"
}

#Create a custom object for every domain in AD Forest and store them in an array
$DomainData = $Forest.Domains | ForEach-Object -Process {
    try
    {
        #creation of the custom object
        New-Object -TypeName PSObject -Property @{
            DomainName = $_.Name
            ParentDomain = $_.Parent
            #Combine all child domains in one string
            ChildDomains = ($_.Children | ForEach-Object {$_.Name}) -join ', '
            DomainMode = $_.DomainMode
            DN = $_.GetDirectoryEntry().distinguishedName[0]
            #Get domain NETBIOS name
            NetBIOSName = Get-DomainNetBIOSName $_.Name -ErrorAction Stop
            PDC = $_.PdcRoleOwner
        }
    }
    catch
    {
        #Display a message in case of an error
        Write-Warning -Message ("Error retrieving domain data {0}: {1}" -f $_.Name, $_.Exception.Message)
    }
}
#Convert Domain info to HTML code
$DomainHTML = $DomainData | ConvertTo-Html -Fragment -As Table -PreContent "<h2>Domains:</h2>" | Out-String
#Create custom objects from Domain info and add them to the xmlData array
$DomainData | %{$xmlData += [PSCustomObject]@{Type="AD";DomainName=$_.DomainName;PDC=$_.PDC;NetBIOSName=$_.NetBIOSName}}

#Get AD Site info and convert to HTML code
$SitesHTML = Get-ADSite | ConvertTo-Html -Fragment -As Table -Property Name, HubSiteEnabled -PreContent "<h2>Sites:</h2>" | Out-String
#Get AD Site Links info and convert to HTML code
$SitesLinkHTML = Get-ADSiteLink | ConvertTo-Html -Fragment -As Table -Property Name, Cost, ADCost, ExchangeCost, @{L='Sites';E={($_.Sites | foreach {$_.Rdn.EscapedName}) -join ', ' }} -PreContent "<h2>Site links:</h2>" | Out-String

#Get Exchange Organization info
$OrgConfigData = Get-OrganizationConfig
#Add Exchange Organization name to the xmlData array
$xmlData += [PSCustomObject]@{Type="Org";Name=$OrgConfigData.Name}
#Convert Exchange Organization info to HTML code
$OrgConfigHTML = $OrgConfigData | ConvertTo-Html -Fragment -As List -Property Name, AdminDisplayName -PreContent "<h2>Organization Info:</h2>" | Out-String

#Resolve Build number to CU friendly name
function Get-ExchangeUpdateName($build)
{
	switch($build)
	{
		#Exchange 2019
		{$build -like "Version 15.2 (Build 221.12)"} {"Exchange 2019 RTM"}
		#Exchange 2016
		{$build -like "Version 15.1 (Build 225.16)"} {"Exchange 2016 RTM"}
		{$build -like "Version 15.1 (Build 396.30)"} {"Exchange 2016 CU1"}
		{$build -like "Version 15.1 (Build 466.34)"} {"Exchange 2016 CU2"}
		{$build -like "Version 15.1 (Build 544.27)"} {"Exchange 2016 CU3"}
		{$build -like "Version 15.1 (Build 669.32)"} {"Exchange 2016 CU4"}
		{$build -like "Version 15.1 (Build 845.34)"} {"Exchange 2016 CU5"}
		{$build -like "Version 15.1 (Build 1034.26)"} {"Exchange 2016 CU6"}
		{$build -like "Version 15.1 (Build 1261.35)"} {"Exchange 2016 CU7"}
		{$build -like "Version 15.1 (Build 1415.2)"} {"Exchange 2016 CU8"}
		{$build -like "Version 15.1 (Build 1466.3)"} {"Exchange 2016 CU9"}
		{$build -like "Version 15.1 (Build 1531.3)"} {"Exchange 2016 CU10"}
		{$build -like "Version 15.1 (Build 1591.01)"} {"Exchange 2016 CU11"}
		#Exchange 2013
		{$build -like "Version 15.0 (Build 516.32)"} {"Exchange 2013 RTM"}
		{$build -like "Version 15.0 (Build 620.29)"} {"Exchange 2013 CU1"}
		{$build -like "Version 15.0 (Build 712.24)"} {"Exchange 2013 CU2"}
		{$build -like "Version 15.0 (Build 775.38)"} {"Exchange 2013 CU3"}
		{$build -like "Version 15.0 (Build 847.32)"} {"Exchange 2013 CU4"}
		{$build -like "Version 15.0 (Build 913.22)"} {"Exchange 2013 CU5"}
		{$build -like "Version 15.0 (Build 995.29)"} {"Exchange 2013 CU6"}
		{$build -like "Version 15.0 (Build 1044.25)"} {"Exchange 2013 CU7"}
		{$build -like "Version 15.0 (Build 1076.9)"} {"Exchange 2013 CU8"}
		{$build -like "Version 15.0 (Build 1104.5)"} {"Exchange 2013 CU9"}
		{$build -like "Version 15.0 (Build 1130.7)"} {"Exchange 2013 CU10"}
		{$build -like "Version 15.0 (Build 1156.6)"} {"Exchange 2013 CU11"}
		{$build -like "Version 15.0 (Build 1178.4)"} {"Exchange 2013 CU12"}
		{$build -like "Version 15.0 (Build 1210.3)"} {"Exchange 2013 CU13"}
		{$build -like "Version 15.0 (Build 1236.3)"} {"Exchange 2013 CU14"}
		{$build -like "Version 15.0 (Build 1263.5)"} {"Exchange 2013 CU15"}
		{$build -like "Version 15.0 (Build 1293.2)"} {"Exchange 2013 CU16"}
		{$build -like "Version 15.0 (Build 1320.4)"} {"Exchange 2013 CU17"}
		{$build -like "Version 15.0 (Build 1347.2)"} {"Exchange 2013 CU18"}
		{$build -like "Version 15.0 (Build 1365.1)"} {"Exchange 2013 CU19"}
		{$build -like "Version 15.0 (Build 1367.3)"} {"Exchange 2013 CU20"}
		{$build -like "Version 15.0 (Build 1395.4)"} {"Exchange 2013 CU21"}
		#Exchange 2010
		{$build -like "Version 14.3 (Build 123.4)"} {"Exchange 2010 SP3"}
		default {"Exchange 20??"}
	}
}

#Get info for DAG that is created last
$DAGData = Get-DatabaseAvailabilityGroup | Sort-Object WhenCreatedUTC -Descending | Select-Object -First 1
#Add DAG info to the xmlData array
$DAGData | %{$xmlData += [PSCustomObject]@{Type="DAG";Name=$_.Name;Servers=$_.Servers.Name}}

#Convert DAG info to HTML code
$DAGsHTML = $DAGData | ConvertTo-Html -Property Name, @{'L'='Servers';'E'={$_.Servers.Name -join ' '}}, WitnessServer, DatacenterActivationMode, @{l='IPv4 Addressess';e={($_.DatabaseAvailabilityGroupIpv4Addresses | Select-Object -ExpandProperty IPAddressToString) -Join ','}}, ThirdpartyReplication, AllowCrossSiteRpcClientAccess -Fragment -As Table -PreContent "<h2>Database Availability Groups:</h2>" | Out-String

#Convert DAG networks for the selected DAG to HTML code
$DAGNetworksHTML = Get-DatabaseAvailabilityGroupNetwork -Identity $DAGData.Name | ConvertTo-Html -Property Name, @{l='Subnets';e={$_.Subnets | % {$_.SubnetId.IPRange.Expression}}}, MapiAccessEnabled,ReplicationEnabled,IgnoreNetwork -Fragment -As Table -PreContent "<h2>DAG Networks:</h2>" | Out-String

#Get Servers for the selected DAG
$DAGServers = $DAGData.Servers | Sort-Object
#Get details for every DAG server
$ExchangeSrvs = ForEach ($DAGSrv in $DAGServers){Get-ExchangeServer -Status $DAGSrv.Name}
#Collect names of DAG members
$DagServerNames = ForEach ($DAGSrv in $ExchangeSrvs) {$DAGSrv.Name} 
#Get Build version for the first server in the selected DAG
$WhichExVer = $ExchangeSrvs[0].AdminDisplayVersion
#Get AD Site for the first server in the selected DAG
$WhichExSite = $ExchangeSrvs[0].Site

#Create custom object for every DAG server
$ExchangeServers = $ExchangeSrvs | ForEach-Object {
    New-Object -TypeName PSObject -Property @{
        ServerName = $_.Name
        Domain = $_.Domain
        Site = $_.Site
        ServerRoles = $_.ServerRole
        #Combine all GC servers in one string
        GC = $_.CurrentGlobalCatalogs -join ', '
        Edition = $_.Edition
        FQDN = $_.Fqdn
        OSVersion = ''
        OSSPVersion = ''
        Disks = ''
        #Get Exchange version
        ExVersion = Get-ExchangeUpdateName($_.AdminDisplayVersion)
        'IPv4 Addresses' = ''
        'Subnet Mask' = ''
        'Default Gateway' = ''
        'DNS Servers' = ''
    }
}

#Start data an asynchronous data collection for every Exchange server
$ExchangeData = $ExchangeServers | where {$_.ServerRoles -notlike "*edge*"} | Invoke-RunspaceJob -ThrottleLimit 50 -Timeout 2 -ScriptBlock {
    #Get Exchange Server object
    $ResultObject = $_
    #Get OS Version from Exchange server's WMI database
    $OS = Get-WmiObject -Class Win32_OperatingSystem -Property Caption, CSDVersion -ComputerName  $ResultObject.Fqdn -ErrorAction Stop
    #Get Volume info from Exchange server's WMI database and convert to strings
    $Disks = (Get-WmiObject -Class Win32_Volume -Property Name, Capacity, FreeSpace, BlockSize -Filter "DriveType = 3" -ComputerName $ResultObject.Fqdn -ErrorAction Stop | ForEach-Object -Process {
        "Path={0}; Capacity={1:N2} GB; Free={2:N2} GB; Cluster={3} KB" -f $_.Name, ($_.Capacity / 1gb), ($_.FreeSpace / 1gb), ($_.BlockSize / 1kb)
    }) -join "---------"
    #Get Network Adapter IP configuration from Exchange server's WMI database
    $IPConfig  = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName  $ResultObject.Fqdn -ErrorAction Stop | ? {$_.DefaultIPGateway -ne $Null} | 
                    Select-Object @{l='IPv4 Addresses';e={$_.IPAddress -match '(\d{1,3}\.){3}\d{1,3}'}}, `
                                  @{l='Subnet Mask';e={@($_.IPSubnet)[0]}}, `
                                  @{l='Default Gateway';e={$_.DefaultIPGateway}}, `
                                  @{l='DNS Servers';e={$_.DNSServerSearchOrder}}
    #Add OS info to the Result object
    if ($OS)
    {
        $ResultObject.OSVersion = $OS.Caption
        $ResultObject.OSSPVersion = $OS.CSDVersion
    }
    #Add Disk info to the Result object
    if ($Disks)
    {
        $ResultObject.Disks = $Disks
    }
    #Add Net Adapter info to the Result object
    if ($IPConfig) {
        $ResultObject.'IPv4 Addresses' = $IPConfig.'IPv4 Addresses' -join ','
        $ResultObject.'Subnet Mask' = $IPConfig.'Subnet Mask'
        $ResultObject.'Default Gateway' = $IPConfig.'Default Gateway'
        $ResultObject.'DNS Servers' = $IPConfig.'DNS Servers' -join ','
    }
    #Send the Result object to the function's output
    Write-Output $ResultObject
}

#Add to xmlData array OS/hardware info for Exchange servers 
$exchangeData | %{$xmlData += [PSCustomObject]@{Type="Server";Name=$_.ServerName;ServerRoles=$_.ServerRoles;FQDN=$_.FQDN;Site=$_.Site;Domain=$_.Domain;GC=$_.GC;Disks=$_.Disks;ExVersion=$_.ExVersion}}
#Prepare HTML code for OS/hardware info
$ExchangeServersHTML = $ExchangeData | sort Site,ServerRoles,ServerName | ConvertTo-Html -Property ServerName,ServerRoles,FQDN,GC,Site,OSVersion,Disks,OSSPVersion,ExVersion,Edition,Domain -Fragment -As Table -PreContent "<h2>Servers:</h2>" | Out-String

#Prepare HTML code for Network Adapter configuration
$TCPIPHTML = $ExchangeData | ConvertTo-Html -Property ServerName, 'IPv4 Addresses', 'Subnet Mask', 'Default Gateway', 'DNS Servers' -Fragment -As Table -PreContent "<h2>TCP/IP Config:</h2>" | Out-String

#Get list of all databases for selected DAG
$MailboxDBs = ForEach ($DAGSrv in $DAGServers){Get-MailboxDatabase -Server $DAGSrv.Name}
$MailboxDBs =  $MailboxDBs | select -Unique

#Prepare Database info objects and add them to the xmlData array
$MailboxData = $MailboxDBs | Sort-Object Name | Select-Object Name, Server, EdbFilePath, LogFolderPath, MasterServerOrAvailabilityGroup, MasterType, Recovery, @{L='Copies';E={($_.DatabaseCopies | Select-Object -ExpandProperty HostServerName) -join ' '}}
$MailboxData | %{$xmlData += [PSCustomObject]@{Type="DB";Name=$_.Name;EdbFilePath=$_.EdbFilePath;LogFolderPath=$_.LogFolderPath;DAG=$_.MasterServerOrAvailabilityGroup;MasterType=$_.MasterType;Recovery=$_.Recovery;Copies=$_.Copies}}

#Prepare HTML code for Database info
$MailboxDBHTML = $MailboxData | ConvertTo-Html -Fragment -As Table -PreContent "<h2>Mailbox Databases:</h2>" | Out-String

#Get Exchange Edge subscription info
$EdgeSubscription = Get-EdgeSubscription
#Prepare HTML code for Edge info in case of an existing Edge subscription
if ($EdgeSubscription -ne $null) {
    $EdgeHTML = $EdgeSubscription | ConvertTo-Html -Fragment -As Table -Property Name, Site, Domain -PreContent "<h2>Edge:</h2>" | Out-String
}

#Get list of accepted email domains
$AcceptedDomainsData = Get-AcceptedDomain
#Add email domains to the xmlData array
$AcceptedDomainsData | ?{$_.Default -eq "True"} | %{$xmlData += [PSCustomObject]@{Type="AcceptedDomains";Name=$_.Name;DomainName=$_.DomainName}}

#Prepare HTML code for accepted email domains
$AcceptedDomainsHTML = $AcceptedDomainsData | ConvertTo-Html -Property Name, DomainName, DomainType, Default -Fragment -As Table -PreContent "<h2>Accepted Domains:</h2>" | Out-String

#Get Email address policy and prepare HTML code  
$EmailAddressPoliciesHTML = Get-EmailAddressPolicy | ConvertTo-Html Name, Priority, EnabledPrimarySMTPAddressTemplate, IncludedRecipients -Fragment -As Table -PreContent "<h2>Email Address Policies:</h2>" | Out-String

#Check Exchange version and get Client Access Server/Service (CAS) config
Switch ($WhichExVer)
	{
        #For Exchange 2010
		{$WhichExVer -like "*14.*"} {
            #Get a list of Client Access Servers (CAS) in a selected site
			$CASServers = Get-ExchangeServer | ? {$_.Site -eq $WhichExSite -and $_.ServerRole -like "*ClientAccess*" -and $_.AdminDisplayVersion -like "*14.*"}
            #Get additional info for CAS servers
			$CASConfig = ForEach ($CASSrv in $CASServers){Get-ClientAccessServer -IncludeAlternateServiceAccountCredentialStatus -Identity $CASSrv.Name | Select-Object Name, AutoDiscoverServiceInternalUri, AlternateServiceAccountConfiguration}
            #Exchange version
            $ExchangeVer = '2010'
            #Exchange major build
            $ExchangeBuild = '14.03'
		}
        #For Exchange 2013
		{$WhichExVer -like "*15.0*"} {
            #Get a list of Client Access Servers (CAS) in a selected site
			$CASServers = Get-ExchangeServer | ? {$_.Site -eq $WhichExSite -and $_.ServerRole -like "*ClientAccess*" -and $_.AdminDisplayVersion -like "*15.0*"}
			$CASConfig = ForEach ($CASSrv in $CASServers){Get-ClientAccessServer -IncludeAlternateServiceAccountCredentialStatus -Identity $CASSrv.Name | Select-Object Name, AutoDiscoverServiceInternalUri, AlternateServiceAccountConfiguration}
            $ExchangeVer = '2013'
            #Exchange major build
            $ExchangeBuild = '15.00'
		}
        #For Exchange 2016
		{$WhichExVer -like "*15.1*"} {
            #Get a list of Client Access Services (CAS)
			$CASServers = $DAGServers
            #Get additional info for CAS node
			$CASConfig = ForEach ($CASSrv in $CASServers) {Get-ClientAccessService -IncludeAlternateServiceAccountCredentialStatus -Identity $CASSrv.Name | Select-Object Name, AutoDiscoverServiceInternalUri, AlternateServiceAccountConfiguration}
            #Exchange version
            $ExchangeVer = '2016'
            #Exchange major build
            $ExchangeBuild = '15.01'
		}
        #For Exchange 2019
		{$WhichExVer -like "*15.2*"} {
            #Get a list of Client Access Services (CAS)
			$CASServers = $DAGServers
            #Get additional info for CAS node
			$CASConfig = ForEach ($CASSrv in $CASServers) {Get-ClientAccessService -IncludeAlternateServiceAccountCredentialStatus -Identity $CASSrv.Name | Select-Object Name, AutoDiscoverServiceInternalUri, AlternateServiceAccountConfiguration}
            #Exchange version
            $ExchangeVer = '2019'
            #Exchange major build
            $ExchangeBuild = '15.02'
		}
	}

#Get a filtered list of names for Client Access Servers in a selected site and populate CAS1Name and CAS2Name variables
$CASNames = $CASServers | ?{$DagServerNames -contains $_.Name} | select -First 2 | %{$_.Name}
    switch ($CASNames.Count) 
    {
        2 {$CAS1Name = $CASNames[0]; $CAS2Name = $CASNames[1]}
        1 {$CAS1Name = $CASNames[0]
           $CAS2Name = ($CASServers | ?{$_.name -ne $CAS1Name} | select -First 1).Name }
  default {
           $CAS1Name = ($CASServers | select -First 1).Name
           $CAS2Name = ($CASServers | ?{$_.name -ne $CAS1Name} | select -First 1).Name
          }
    }

#Add to the xmlData array CAS info
$xmlData += [PSCustomObject]@{Type="CASInfo";CAS1Name=$CAS1Name;CAS2Name=$CAS2Name}
#Add to the xmlData array Exchange version info
$xmlData += [PSCustomObject]@{Type="GeneralInfo";ExchangeVer=$ExchangeVer;ExchangeBuild=$ExchangeBuild}

#Prepare HTML code for CAS info
$CASConfigHTML = $CASConfig | ConvertTo-Html -Fragment -As Table -PreContent "<h2>CAS Config:</h2>" | Out-String

#Prepare an empty array for URL and authentication config of Virtual Directories
$URLConfig = @()
#Get Outlook Anywhere config and add it to URLConfig array
$URLConfig += ForEach ($CASSrv in $CASServers){Get-OutlookAnywhere -Server $CASSrv.Name -ADPropertiesOnly | Select-Object Identity, @{L='InternalUrl';E={$_.InternalHostName}}, @{L='ExternalUrl';E={$_.ExternalHostName}}, @{L='InternalAuthenticationMethods';E={$_.InternalClientAuthenticationMethod}}, @{L='ExternalAuthenticationMethods';E={$_.ExternalClientAuthenticationMethod}}, @{L='IISAuthenticationMethods';E={($_.IISAuthenticationMethods) -join ' '}}}
#Get MAPI over HTTP config and add it to URLConfig array
If ($WhichExVer -like "*15.*") {$URLConfig += ForEach ($CASSrv in $CASServers){Get-MAPIVirtualDirectory -Server $CASSrv.Name -ADPropertiesOnly | Select-Object Identity, InternalUrl, ExternalUrl, @{L='InternalAuthenticationMethods';E={($_.InternalAuthenticationMethods) -join ' '}}, @{L='ExternalAuthenticationMethods';E={($_.ExternalAuthenticationMethods) -join ' '}}, @{L='IISAuthenticationMethods';E={($_.IISAuthenticationMethods) -join ' '}}}}
#Get Offline Address Book (OAB) config and add it to URLConfig array
$URLConfig += ForEach ($CASSrv in $CASServers){Get-OABVirtualDirectory -Server $CASSrv.Name -ADPropertiesOnly | Select-Object Identity, InternalUrl, ExternalUrl, @{L='InternalAuthenticationMethods';E={($_.InternalAuthenticationMethods) -join ' '}}, @{L='ExternalAuthenticationMethods';E={($_.ExternalAuthenticationMethods) -join ' '}}}
#Get Exchange Web Services (EWS) config and add it to URLConfig array
$URLConfig += ForEach ($CASSrv in $CASServers){Get-WebServicesVirtualDirectory -Server $CASSrv.Name -ADPropertiesOnly | Select-Object Identity, InternalUrl, ExternalUrl, @{L='InternalAuthenticationMethods';E={($_.InternalAuthenticationMethods) -join ' '}}, @{L='ExternalAuthenticationMethods';E={($_.ExternalAuthenticationMethods) -join ' '}}}
#Get Outlook Web App (OWA) config and add it to URLConfig array
$URLConfig += ForEach ($CASSrv in $CASServers){Get-OwaVirtualDirectory -Server $CASSrv.Name -ADPropertiesOnly | Select-Object Identity, InternalUrl, ExternalUrl, @{L='InternalAuthenticationMethods';E={($_.InternalAuthenticationMethods) -join ' '}}, @{L='ExternalAuthenticationMethods';E={($_.ExternalAuthenticationMethods) -join ' '}}}
#Get Exchange Control Panel (ECP) config and add it to URLConfig array
$URLConfig += ForEach ($CASSrv in $CASServers){Get-EcpVirtualDirectory -Server $CASSrv.Name -ADPropertiesOnly | Select-Object Identity, InternalUrl, ExternalUrl, @{L='InternalAuthenticationMethods';E={($_.InternalAuthenticationMethods) -join ' '}}, @{L='ExternalAuthenticationMethods';E={($_.ExternalAuthenticationMethods) -join ' '}}}
#Get ActiveSync config and add it to URLConfig array
$URLConfig += ForEach ($CASSrv in $CASServers){Get-ActiveSyncVirtualDirectory -Server $CASSrv.Name -ADPropertiesOnly | Select-Object Identity, InternalUrl, ExternalUrl, @{L='InternalAuthenticationMethods';E={($_.InternalAuthenticationMethods) -join ' '}}, @{L='ExternalAuthenticationMethods';E={($_.ExternalAuthenticationMethods) -join ' '}}}
#Get Autodiscover virtual directory config and add it to URLConfig array
$URLConfig += ForEach ($CASSrv in $CASServers){Get-AutodiscoverVirtualDirectory -Server $CASSrv.Name -ADPropertiesOnly | Select-Object Identity, InternalUrl, ExternalUrl, @{L='InternalAuthenticationMethods';E={($_.InternalAuthenticationMethods) -join ' '}}, @{L='ExternalAuthenticationMethods';E={($_.ExternalAuthenticationMethods) -join ' '}}}
#Prepare HTML code for URLConfig array
$URLConfigHTML = $URLConfig | ConvertTo-Html -Fragment -As Table -PreContent "<h2>CAS URLs and Authentication:</h2>" | Out-String

#Combine all HTML parts and save result to ExRESScoping.html file
$Header + "<h1>Exchange Recovery Execution Service (ExRES) Scoping Tool</h1>" + $ForestHTML + $OrgConfigHTML + $DomainHTML + $SitesHTML + $SitesLinkHTML + $DAGsHTML + $ExchangeServersHTML + $EdgeHTML + $TCPIPHTML + $DAGNetworksHTML + $MailboxDBHTML + $AcceptedDomainsHTML + $EmailAddressPoliciesHTML + $CASConfigHTML + $URLConfigHTML | Out-File $HTMLOutPath -Force
#Save collected objects in xmlData array to xml file
$xmlData | Export-Clixml -Path $XMLOutPath
#Archive ExRESScoping.html and ExRESScoping.xml files to ExRESScoping.zip archive
Create-ZipFile -dir $MyDocsPath
#Get full path to ExRESScoping.zip file
$zipPath = Join-Path -Path $MyDocsPath -ChildPath "ExRESScoping.zip"
#Start Windows Explorer and open directory with ExRESScoping.zip file
Invoke-Expression "explorer.exe '/select,$zipPath'"
#Display a reminder in the Powershell console
Write-Host "Please upload the $zipPath file to the secure UDE Workspace provided by the MS Engineer."
