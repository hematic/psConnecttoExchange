Function New-ExchangeConnection{
    <#
        .SYNOPSIS
        Function to make a remote connection to an exchange cas server and import commandlets.
        .DESCRIPTION
        This function allows you to specify required settings and then picks the best CAS server that fits those settings and imports a pssession.
        As of version 1.2.5 it properly works with PS 5.1 and PS 7
        .EXAMPLE
        $Splat = @{
            Credential = $Credential
            ConnectionURI = 'http://servername.yourdomain.com/PowerShell'
        }
        New-ExchangeConnection @Splat
        .EXAMPLE
        $Splat = @{
            Credential = $Credential
            Region = 'EM1'
        }
        New-ExchangeConnection @Splat
        .EXAMPLE
        $Splat = @{
            Credential = $Credential
            Region = 'AP1'
            Version = 'Exchange 2016'
        }
        New-ExchangeConnection @Splat
        .EXAMPLE
        $Splat = @{
            Credential = $Credential
            Region = 'EM1'
            Version = 'Exchange 2019'
        }
        New-ExchangeConnection @Splat
    #>
    [CmdletBinding(SupportsShouldProcess=$False,ConfirmImpact='Low')]
    Param(
        #Credential parameter is mandatory in ALL parameter sets.
        [Parameter(Mandatory=$True,ParameterSetName = 'ByServer')]
        [Parameter(Mandatory=$True,ParameterSetName = 'ByRegion')]
        [Parameter(Mandatory=$True,ParameterSetName = 'ByRegionandVersion')]
        
        [Parameter(Mandatory=$True,HelpMessage='This is credential for connecting to Exchange.')]
        [ValidateNotNullOrEmpty()]
        [PSCredential]$Credential,

        #ConnectionURI is used only in the ByServer parameter set
        [Parameter(Mandatory=$True,ParameterSetName = 'ByServer')]

        [Parameter(Mandatory=$False, HelpMessage='This is the CAS Server to connect to.')]
        [ValidateNotNullOrEmpty()]
        [String]$ConnectionURI,

        #Region is used in both the ByRegion and the ByRegionandVersion parameter sets
        [Parameter(Mandatory=$True,ParameterSetName = 'ByRegion')]
        [Parameter(Mandatory=$True,ParameterSetName = 'ByRegionandVersion')]

        [Parameter(Mandatory=$False, HelpMessage='This is the region of the server you want.')]
        [ValidateSet('AM1','EM1','AP1')]
        [String]$Region,

        #Version is used in the ByRegionandVersion parameter set
        [Parameter(Mandatory=$True,ParameterSetName = 'ByRegionandVersion')]

        [Parameter(Mandatory=$False, HelpMessage='This is the exchange version required.')]
        [ValidateSet(,'Exchange 2016','Exchange 2019')]
        [String]$Version
    )

    $Servers = @(
                 [pscustomobject]@{name ="<exchange_server_name>.<domain>.com"	;Version = "Exchange 2016"; Region = "AM1"}
				 [pscustomobject]@{name ="<exchange_server_name>.<domain>.com"	;Version = "Exchange 2016"; Region = "AM1"}
				 [pscustomobject]@{name ="<exchange_server_name>.<domain>.com"	;Version = "Exchange 2016"; Region = "EM1"}
				 [pscustomobject]@{name ="<exchange_server_name>.<domain>.com"	;Version = "Exchange 2016"; Region = "EM1"}
				 [pscustomobject]@{name ="<exchange_server_name>.<domain>.com"	;Version = "Exchange 2016"; Region = "AP1"}
				 [pscustomobject]@{name ="<exchange_server_name>.<domain>.com"	;Version = "Exchange 2016"; Region = "AP1"}
                 [pscustomobject]@{name ="<exchange_server_name>.<domain>.com"	;Version = "Exchange 2019"; Region = "AM1"}
                 [pscustomobject]@{name ="<exchange_server_name>.<domain>.com"  ;Version = "Exchange 2019"; Region = "AM1"}
                 [pscustomobject]@{name ="<exchange_server_name>.<domain>.com"	;Version = "Exchange 2019"; Region = "EM1"}
                 [pscustomobject]@{name ="<exchange_server_name>.<domain>.com"	;Version = "Exchange 2019"; Region = "EM1"}
                 [pscustomobject]@{name ="<exchange_server_name>.<domain>.com"	;Version = "Exchange 2019"; Region = "AM1"}
                 [pscustomobject]@{name ="<exchange_server_name>.<domain>.com"	;Version = "Exchange 2019"; Region = "AP1"}

    )
    
    switch ($PsCmdlet.ParameterSetName){
        'ByServer' {
            #No Specific Processing Needed for this ParameterSet
        }

        'ByRegion'{
            Write-Host "Finding the Fastest server in $Region"
            #Filter the list of servers to limit the number we test against.
            $Servers = $Servers | Where-Object {$_.region -eq $Region}
            $Servers.name | ForEach-Object {
                Start-Job -ScriptBlock {
                    param($Server)
                    Test-Connection -ComputerName $Server
                } -ArgumentList $_
            }
        }
        'ByRegionandVersion'{
            Write-Host "Finding the Fastest $Version server in $Region"
            $Servers = $Servers | Where-Object {$_.Region -eq $Region -and $_.Version -eq $Version}
            $Servers.name | ForEach-Object {
                Start-Job -ScriptBlock {
                    param($Server)
                    Test-Connection -ComputerName $Server
                } -ArgumentList $_
            }
        }
    }
    Switch ($PSVersionTable.PSVersion.Major){
        '5' {
            Get-Job | Wait-Job
            $ConnectionServer = (Get-Job | Receive-Job | Sort-Object -Property ResponseTime | Select-Object -First 1).Address            
        }
        '7' {
            Get-Job | Wait-Job
            $ConnectionServer = (Get-Job | Receive-Job | Sort-Object -Property ResponseTime | Select-Object -First 1).Destination
        }
        default{Write-Error "Neither PS 5 nor PS 7 were detected."}
    }
    If(!$ConnectionServer){
        Write-Error "No Server was able to meet the given criteria."
    }
    $ConnectionURI = 'http://' + $ConnectionServer + '/PowerShell/'
    #Was experiencing an error here that i find at the below link as well
    #https://social.technet.microsoft.com/Forums/lync/en-US/72d3fa90-e134-4bee-aa12-2e190da8bf7c/importpssession-attribute-cannot-be-added-because-it-would-cause-the-variable?forum=winserverpowershell
    #Solution is to remove these extra variables before i do the import-pssesion
    Remove-Variable Region
    Remove-Variable Version
    
    Try{
        $NewSessionSplat = @{
            ConfigurationName = 'microsoft.exchange'
            ConnectionUri = $ConnectionURI
            Authentication = 'Kerberos'
            Credential = $Credential
            SessionOption = New-PSSessionOption -IdleTimeout 7000000 -NoCompression
            ErrorAction = 'Stop'
        }
        $ExchangeSession = New-PSSession @NewSessionSplat
        $ImportSessionSplat = @{
            Session = $ExchangeSession
            ErrorAction = 'Stop'
            WarningAction = 'SilentlyContinue'
            AllowClobber = $True
            DisableNameChecking = $True
        }
        Import-Module (Import-PSSession @ImportSessionSplat) -Global -DisableNameChecking
    }
    Catch{
        Write-Error $_
    }
}