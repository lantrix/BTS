#
# Create a new BizTalk host
#
function CreateHost([string]$Name, [string]$HostType, [boolean]$AuthTrusted, [string]$NtGroupName, [boolean]$Is32BitOnly, [boolean]$IsTrackingHost, [boolean]$IsDefault)
{
	# Set location to add host
    Set-Location 'BizTalk:\Platform Settings\Hosts'
 
    # Check if the host exists
    if (test-path $Name)
    {   
        # Host allready exists
        write-host $Name ' allready exists'
    }
    else
    {
        # Create the host
        New-Item -Path:$Name -HostType:$HostType -NtGroupName:$NtGroupName -AuthTrusted:$AuthTrusted
 
		# Set if this is the tracking host
        Set-ItemProperty -Path:$Name -Name:HostTracking -Value:$IsTrackingHost
 
		# Set is the host is 32 bit only
        Set-ItemProperty -Path:$Name -Name:Is32BitOnly -Value:$Is32BitOnly
 
		# Can not set default to false
		if($IsDefault)
		{
			# Set if this is the default host
			Set-ItemProperty -Path:$Name -Name:IsDefault -Value:$IsDefault
		}
    }
}
#
# Create a new BizTalk host instance
#
function CreateHostInstance([string]$HostName, [string]$AccountName, [string]$Password)
{   
    # Set location to add host instance
    Set-Location 'BizTalk:\Platform Settings\Host Instances'
 
	# Create the name to use when checking if the host instance exists
    $HostNameToCheckForExistance = '*' + $HostName + '*'
 
    # Check if host instance exists
    if (test-path $HostNameToCheckForExistance)
    {   
        # Host instance allready exists
        write-host 'Host instance for ' $HostName ' allready exists'
    }
    else
    {   
		# Secure the password before passing it on
        $SecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force
 
		# Set the credentials under which the host instance should run
        $Credentials = New-Object -TypeName:System.Management.Automation.PSCredential -ArgumentList $AccountName, $SecurePassword
 
        # Create the host instance
        New-Item -Path:'HostInstance' -HostName:$HostName -RunningServer:$Env:ComputerName -Credentials:$Credentials
 
		# Start the host instance
		Start-HostInstance -Path:$HostNameToCheckForExistance
    }
}

#
# Create handlers for the adapters on the new hosts, so these get used for processing.
#
function CreateHandler([string]$AdapterName, [string]$HostName, [string]$Direction, [string]$OriginalHost, [boolean]$Default)
{
    # Set location for adapters
    Set-Location 'BizTalk:\Platform Settings\Adapters\'
 
	# Go to the current adapter
    cd $AdapterName
 
	# Create the name to use when checking if the handler for the adapter exists
    $NameToCheckForExistance = '*' + $HostName + '*'
 
	# The name of the old handler
    $OldHandler = $AdapterName + ' ' + $Direction + ' Handler (' + $OriginalHost + ')'
 
    # Check if handler exists for the host
    if (test-path $NameToCheckForExistance)
    {   
        # Handler allready exists
        write-host 'Handler for ' $HostName ' allready exists'
    }
    else
    {   
		# Check on the host name
        switch ($HostName)
        {
            'HTTPSendHost'
            {
                # Because of missing properties (RequestTimeout) existing handler will be changed.
                set-itemproperty -path $OldHandler -Name HostName -Value $HostName
            }
            'SOAPSendHost'
            {
                # Because of missing properties (UseProxy) existing handler will be changed.
                set-itemproperty -path $OldHandler -Name HostName -Value $HostName
            }
            'SMTPSendHost'
            {
                # Because of missing properties (SMTPAuthenticate) existing handler will be changed.
                set-itemproperty -path $OldHandler -Name HostName -Value $HostName
            }
            default
            {
                # Create handler
                if($Direction -eq 'send')
                {
					# Create new send handler
                    new-item -path '.\Dummy' -hostname $HostName -direction $Direction -Default:$Default
                }
                else
                {   
					# Create new receive handler
                    new-item -path '.\Dummy' -hostname $HostName -direction $Direction
                }
            }
        }       
    }
 
    # Write to console
    Write-Host 'Checking if old handler' $OldHandler 'exists'
 
	# Check if old handler exists
    if (test-path $OldHandler)
    {   
		# If so, remove the old handler
        remove-item $OldHandler
 
		# Write to console
        write-host 'Old handler for' $HostName 'removed'
    }
}
#
# Create the default hosts and host instances and add them host instances to the various handlers.
#
function CreateDefaultHostInstances
{
	# The names of the hosts
	[string]$ReceiveHostName = 'ReceiveHost'
	[string]$SendHostName = 'SendHost'
	[string]$OrchestrationsHostName = 'OrchestrationsHost'
	[string]$TrackingHostName = 'TrackingHost'
 
	# The name of the old host, that is used by default by the adapters
	[string]$OldHostName = 'BizTalkServerApplication'
 
	# Create a host for receiving
	CreateHost $ReceiveHostName 'InProcess' $false $NtGroupName $false $false $false
 
	# Create a host instance for receiving
	CreateHostInstance $ReceiveHostName $AccountName $Password
 
	# Set adapters that should be handled by this host instance
	CreateHandler 'FILE' $ReceiveHostName 'Receive' $OldHostName $false
	CreateHandler 'FTP' $ReceiveHostName 'Receive' $OldHostName $false
	CreateHandler 'MQSeries' $ReceiveHostName 'Receive' $OldHostName $false
	CreateHandler 'MSMQ' $ReceiveHostName 'Receive' $OldHostName $false
	CreateHandler 'POP3' $ReceiveHostName 'Receive' $OldHostName $false
	CreateHandler 'SOAP' $SendHostName 'Receive' $OldHostName $false
	CreateHandler 'SQL' $ReceiveHostName 'Receive' $OldHostName $false
	CreateHandler 'WCF-Custom' $ReceiveHostName 'Receive' $OldHostName $false
	CreateHandler 'WCF-NetMsmq' $ReceiveHostName 'Receive' $OldHostName $false
	CreateHandler 'WCF-NetNamedPipe' $ReceiveHostName 'Receive' $OldHostName $false
	CreateHandler 'WCF-NetTcp' $ReceiveHostName 'Receive' $OldHostName $false
	CreateHandler 'Windows SharePoint Services' $ReceiveHostName 'Receive' $OldHostName $false
 
	# Create a host for sending
	CreateHost $SendHostName 'InProcess' $false $NtGroupName $false $false $false
 
	# Create a host instance for sending
	CreateHostInstance $SendHostName $AccountName $Password
 
	# Set adapters that should be handled by this host instance
	CreateHandler 'FILE' $SendHostName 'Send' $OldHostName $true
	CreateHandler 'FTP' $SendHostName 'Send' $OldHostName $true
	CreateHandler 'HTTP' $SendHostName 'Send' $OldHostName $true
	CreateHandler 'MQSeries' $SendHostName 'Send' $OldHostName $true
	CreateHandler 'MSMQ' $SendHostName 'Send' $OldHostName $true
	CreateHandler 'SMTP' $SendHostName 'Send' $OldHostName $true
	CreateHandler 'SOAP' $SendHostName 'Send' $OldHostName $true
	CreateHandler 'SQL' $SendHostName 'Send' $OldHostName $true
	CreateHandler 'WCF-BasicHttp' $SendHostName 'Send' $OldHostName $true
	CreateHandler 'WCF-Custom' $SendHostName 'Send' $OldHostName $true
	CreateHandler 'WCF-NetMsmq' $SendHostName 'Send' $OldHostName $true
	CreateHandler 'WCF-NetNamedPipe' $SendHostName 'Send' $OldHostName $true
	CreateHandler 'WCF-NetTcp' $SendHostName 'Send' $OldHostName $true
	CreateHandler 'WCF-WSHttp' $SendHostName 'Send' $OldHostName $true
	CreateHandler 'Windows SharePoint Services' $SendHostName 'Send' $OldHostName $true
 
	# Create a host for orchestrations
	CreateHost $OrchestrationsHostName 'InProcess' $false $NtGroupName $false $false $false
 
	# Create a host instance for orchestrations
	CreateHostInstance $OrchestrationsHostName $AccountName $Password
 
	# Create a host for tracking
	CreateHost $TrackingHostName 'InProcess' $false $NtGroupName $false $true $false
 
	# Create a host instance for tracking
	CreateHostInstance $TrackingHostName $AccountName $Password
}
#
# Check if there are currently any deployed applications
# 
function CheckDeployedApplications
{
	# Set prompt color
	[console]::ForegroundColor = "Green"
 
	# Check with user
	Write-Host("We will not be able to change the host instances if there are any receiveports, sendports or orchestrations deployed")
	$ApplicationsDeployed = Read-Host -prompt ("Are there any of these deployed applications? Y / N")
 
	# Set prompt color
	[console]::ForegroundColor = "Gray"
 
	# Check user answer
	switch($ApplicationsDeployed)
	{
		"Y"
		{
			Write-Host("First undeploy all these applications, then run this script again") -Fore White
			return $true
		}
		"N"
		{
			return $false
		}
		default
		{
			Write-Error " Please provide a valid answer!"
			return CheckDeployedApplications
		}
	}
}

################################ Set your own properties here ##################################
 
# Name of the group the BizTalk hosts should run under
[string]$NtGroupName='BizTalk Application Users'
 
# The username under which the host instances should run
$AccountName='svc_BTSHost'
 
# The password of the user under which the host instances should run
$Password='Albron@123'
 
################################ Done ##########################################################

# Check if there are currently any deployed applications
if(CheckDeployedApplications -eq $false)
{
	# Write to prompt
	Write-Host("Creating hosts and host instances...") -Fore DarkGreen
 
	# Create default hosts, host instances and handlers
	CreateDefaultHostInstances
}

# Do not close until the user presses a key
Write-Host("Press any key to exit...") -Fore White
$null = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")