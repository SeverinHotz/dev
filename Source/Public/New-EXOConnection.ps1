<# 
    .Synopsis
    Cmdlet to connect to exchange online.

    .Description
    Cmdlet to connect to exchange online with modern authentication and proxy settings.
    Proxy settings configured in IE are used by default.
    Please remove your connection before closing the Powershell session.
    Remove-PSSession $ExchangeOnlineSession
    
    .Example
    # Connect to exchange online
    New-SREXOConnection

    .Example
    # Conenct to exchange online using a username
    New-SREXOConnection -UserPrincipalName prename_lastname@domain.com

    .Example
    # Connect to exchange online without using proxy settings.
    New-SREXOConnection -NoProxy

    .Example
    # Please remove your connection before closing
    Remove-PSSession $ExchangeOnlineSession

#>

function New-EXOConnection {

     [CmdletBinding()]
    
         param (
            # Parameter to switch off proxy configuration
            [switch] $NoProxy,    
            # User Principal Name or email address of the user
            [string] $UserPrincipalName = ''
        )
    
        begin {           
        }  
        
        process {

            # Fetch proxy settings from internet explorer and connect
            if (!$NoProxy)
            {
                # Verbose output regarding proxy settings
                write-verbose "Proxy seetting from internet explorer are used"
                $proxysettings = New-PSSessionOption -ProxyAccessType IEConfig
            }
            else
            {
                # Verbose output regarding proxy settings
                write-verbose "No Proxy seetting are used"
            }
                        
            # Create connection to exchange online        
            $null = import-module (Connect-EXOPSSession -PSSessionOption $proxysettings -Userprincipalname $UserPrincipalName) -global -DisableNameChecking

            # Store powershell session in a variable for later use
            $ExchangeOnlineSession = (Get-PSSession | Where-Object { ($_.ConfigurationName -eq 'Microsoft.Exchange') -and ($_.State -eq 'Opened') })[0]
                
        }
        
        end {
        } 

    }