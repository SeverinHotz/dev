<# 
    .Synopsis
    Cmdlet to connect to exchange online.

    .Description
    Cmdlet to connect to exchange online with modern authentication and proxy settings.
    Proxy settings configured in IE are used by default.
    The connection is stored in a variable exchangeOnlineSession for later use.
    
    .Example
    # Connect to exchange online
    New-EXOConnection

    .Example
    # Connect to exchange online without using proxy settings.
    New-EXOConnection -NoProxy

#>

function New-EXOConnection {

     [CmdletBinding()]
    
         param (
            [switch] $NoProxy    
        )
    
        begin {

            # Fetch the path of the exchange online module
            $ModuleName = $PSModuleRoot + "\Private\Microsoft.Exchange.Management.ExoPowershellModule.dll"
            # Import the module
            Import-Module -FullyQualifiedName $ModuleName -Force
            # Create connection to exchange online
            if($NoProxy)
            {
                # connect without proxy server settings
                $null = import-module (Connect-EXOPSSession) -global -DisableNameChecking
            }
            else
            {
                # Fetch proxy settings from internet explorer and connect
                $proxysettings = New-PSSessionOption -ProxyAccessType IEConfig
                $null = import-module (Connect-EXOPSSession -PSSessionOption $proxysettings) -global -DisableNameChecking
            }    
            
        }  
        
        process {
        
        }
        
        end {
        } 

    }