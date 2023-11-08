# psConnecttoExchange
This module allows you to specify required settings and then picks the best CAS server that fits those settings and imports a pssession.
This allows you to run all exchange commands from you own PC.

As of version 1.2.5 it properly works with PS 5.1 and PS 7

## Please note this module is sanitized from my environment

- You will need to open the .psm1 file and place in your own servers in the region below to make this work.

```
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
```

# Current Release Notes (1.2.5)
- 'Updated for Exchange 2019. Also now supports both PS 5.1 and PS 7.0'


# Description

Tired of connecting to a CAS server or exchange server to run commands? Want to easily run exchange commands in larger scripts without being on a CAS server?

This module has you covered! Integrate this into larger projects like onboarding workflows, or use it to run daily exchange tasks from the comfort of your own pc.

Current Release Version [1.2.5](https://www.powershellgallery.com/packages/psConnecttoExchange/1.2.5)

# Basic Usage

Use a specific CAS server.
```
$Splat = @{
    Credential = $Credential
    ConnectionURI = 'http://servername.yourdomain.com/PowerShell'
}
New-ExchangeConnection @Splat
```
Have the function choose the best server in a region.
```
$Splat = @{
    Credential = $Credential
    Region = 'EM1'
}
New-ExchangeConnection @Splat
```
Have the function choose the best server by region and Exchange version.
```
$Splat = @{
    Credential = $Credential
    Region = 'AP1'
    Version = 'Exchange 2016'
}
New-ExchangeConnection @Splat
```

## Installation
```
Install-Module -Name psConnecttoExchange -Repository PSGallery -Force -Scope CurrentUser
```
