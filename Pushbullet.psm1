# Original code by Patrck Lambert, and can be found at https://github.com/dendory/Pushbullet-PowerShell-Module
# Modified and extended by Kevin Doblosky

# Location on the local computer where cached PushBullet Devices and ContactEmails should be stored
$StorageLocation = "E:\Scripts\Settings"

# API Key to use as a default for each function, if none is provided.
$MyAPIKey = ""

Function Send-Pushbullet
{
    <#
    .SYNOPSIS
        Send-Pushbullet can be used with the Pushbullet service to send notifications to your devices.

    .DESCRIPTION
        This function requires an account at Pushbullet. Register at http://pushbullet.com and obtain your API Key from the account settings section.

        With this module you can send messages or links from a remote system to all of your devices.
   
    .EXAMPLE
        Send-Pushbullet -APIKey "XXXXXX" -Title "Hello World" -Message "This is a test of the notification service."

        Send a message to all your devices.

    .EXAMPLE
        Send-Pushbullet -APIKey "XXXXXX" -Title "Here is a link" -Link "http://pushbullet.com" -DeviceIden "XXXXXX"

        Send a link to one of your deivces. Use Get-PushbulletDevices to get a list of Iden codes.

    .EXAMPLE
        Send-Pushbullet -APIKey "XXXXXX" -Title "Hey there" -Message "Are you here?" -ContactEmail "user@example.com"

        Send a message to a remote user.
    #>
    param([string]$APIKey=$MyAPIKey, [string]$Message="", [string]$Link="", [string]$DeviceIden="", [string]$ContactEmail="", [string]$Title="")

    if($Link -ne "")
    {
        $Body = @{
            type = "link"
            title = $Title
            body = $Message
            url = $Link
            device_iden = $DeviceIden
            email = $ContactEmail
        }
    }
    else
    {
        $Body = @{
            type = "note"
            title = $Title
            body = $Message
            device_iden = $DeviceIden
            email = $ContactEmail
        }
    }

    $Creds = New-Object System.Management.Automation.PSCredential ($APIKey, (ConvertTo-SecureString $APIKey -AsPlainText -Force))
    $response = Invoke-WebRequest -Uri "https://api.pushbullet.com/v2/pushes" -Credential $Creds -Method Post -Body $Body
}


Function Send-PushBulletEx{ 
    <#
        .Synopsis
        Send a PushBullet.
        .Description
        Send a PushBullet with autocomplete for devices and email contacts, as well as accepting pipeline input for the
        message. Objects will get piped to Out-String before sending the message.
        .Example
        $> 1..10 | Send-PushBulletEx -DeviceIden MyComputer

        # Sends a single message with the values 1 through 10, 1 on each line.
        .Example
        $> 1..10 | Send-PushBulletEx -DeviceIden MyComputer -SendMultiplePushes

        # Sends a message for each item in the pipeline. So this will send 10 pushes, one for each number.

        .Parameter Message
        Message to send. Can come from the pipeline. If it is an array, how it is processed depends on 
        the value of SendMultiplePushes. If SendMultiplePushes is true, each item is sent as a separate message. 
        Otherwise, all values are collected, piped through Out-String, and then sent as a single message.
        .Parameter APIKey
        APIKey to use for PushBullet. Defaults to using $MyAPIKey, which is specified at the top of this module
        .Parameter Link
        Link to send via PushBullet
        .Parameter Title
        Title of the push to be sent
        .Parameter SendMultiplePushes
        When used with an array, sends an individual push for each array element. Defaults to sending a single push.
        .Parameter DeviceIden
        Dynamic parameter that will provide a dropdown / autocomplete for the valid devices. 
        .Parameter ContactEmail
        Dynamic parameter that will provide a dropdown / autocomplete for the valid contact emails. 
        .Outputs
        Nothing is sent to the pipeline. The result is one or more pushes being sent.
        .Notes
        NAME:  Send-PushbulletEx
        AUTHOR: Kevin Doblosky
        LASTEDIT: 2015-02-13
    #>
    [CmdletBinding()]
    Param(
        [parameter(ValueFromPipeline=$true)]$Message = "",
        [string]$APIKey = $MyAPIKey, 
        [string]$Link = "",
        [string]$Title = "",
        [switch]$SendMultiplePushes
    )

    DynamicParam{
        # Add dynamic parameters for DeviceIden and ContactEmail with ValidateSet of valid options

        $DevicePath = Join-Path $StorageLocation "PushBulletDevices.json"
        $EmailPath = Join-Path $StorageLocation "PushBulletContacts.json"

        # Make sure that we have already cached the devices and contact emails locally. If not, update them
        If (-Not (Test-Path $DevicePath) -Or -Not (Test-Path $EmailPath) ) {
            Update-PushBulletSettings -APIKey $APIKey
        }

        # Add dynamic parameter for DeviceIden
        $attrib = New-Object System.Management.Automation.ParameterAttribute
        $attrib.ParameterSetName = "__AllParameterSets"
        $attrib.Mandatory = $false

        $devices = Get-Content (Join-Path $StorageLocation "PushBulletDevices.json") -Raw | ConvertFrom-Json 
        $items = $devices | Select-Object -ExpandProperty nickname
            
        $validate = New-Object System.Management.Automation.ValidateSetAttribute -ArgumentList $items
        
        $AttributeCollection = New-Object 'Collections.ObjectModel.Collection[System.Attribute]'
        $AttributeCollection.Add($attrib)
        $AttributeCollection.Add($validate)

        $DynParameter = New-Object -TypeName System.Management.Automation.RuntimeDefinedParameter -ArgumentList @("DeviceIden", [string], $AttributeCollection)
        
        $ParamDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
        $ParamDictionary.Add("DeviceIden", $DynParameter)

        # Add dynamic parameter for ContactEmail
        $attrib2 = New-Object System.Management.Automation.ParameterAttribute
        $attrib2.ParameterSetName = "__AllParameterSets"
        $attrib2.Mandatory = $false
        
        $emails = Get-Content (Join-Path $StorageLocation "PushBulletContacts.json") -Raw | ConvertFrom-Json
        $items2 = $emails | Select-Object -ExpandProperty email

        $validate2 = New-Object System.Management.Automation.ValidateSetAttribute -ArgumentList $items2
        
        $AttributeCollection2 = New-Object 'Collections.ObjectModel.Collection[System.Attribute]'
        $AttributeCollection2.Add($attrib2)
        $AttributeCollection2.Add($validate2)

        $DynParameter2 = New-Object -TypeName System.Management.Automation.RuntimeDefinedParameter -ArgumentList @("ContactEmail", [string], $AttributeCollection2)
        
        $ParamDictionary.Add("ContactEmail", $DynParameter2)
        

        return $ParamDictionary
    }

    Begin {
        # Parameters to be used to send push
        $pushParams = @{
            APIKey = $APIKey;
            Link = $Link;
            Title = $Title;
        }

        # DeviceIden is specified using the nickname. Find the iden that corresponds to that nickname
        If (-Not [String]::IsNullOrEmpty($DeviceIden)) {
            $devices = Get-Content (Join-Path $StorageLocation "PushBulletDevices.json") -Raw | ConvertFrom-Json 
            $pushParams["DeviceIden"] = $devices | Where-Object { $_.nickname = $DeviceIden } | 
                Select-Object -ExpandProperty iden
        }

        If (-Not [String]::IsNullOrEmpty($ContactEmail)) {
            $pushParams["ContactEmail"] = $ContactEmail
        }
      

        $results = @()
    }

    Process {
        
        If ($SendMultiplePushes) {

            # Arrays not sent through the pipeline only execute the Process block once.
            # So if we have an array, send a push for each element.
            If ($Message -is [Object[]]) {
                $Message | ForEach-Object {
                    $sentMessage = If ($_ -is [string] ) {
                        $_ 
                    } Else {
                        $_ | Out-String
                    }   
            
                    $pushParams["Message"] = $sentMessage;    
                    Send-Pushbullet @pushParams
                }
            } Else {

                # If the message is already a string, send it as is, otherwise pipe it through Out-String
                $sentMessage = If ($Message -is [string] ) {
                    $Message 
                } else {
                    $Message | Out-String
                }   
            
                $pushParams["Message"] = $sentMessage;    
                Send-Pushbullet @pushParams
            }
            
        } Else {
            # If we're sending a single push, accumulate the results to be sent later
            $results += $Message
        }
    }

    End {
        # If we're sending a single push, send it now.
        If (-Not $SendMultiplePushes) {
            $pushParams["Message"] = $results | Out-String
            Send-Pushbullet @pushParams   
        }
    }
}

Function Update-PushBulletSettings
{
    <#
        .Synopsis
        Updates locally-cached PushBullet devices and contact emails.
        .Description
        Updates locally-cached PushBullet devices and contact emails.
        .Example
        $> Update-PushBulletSettings

        .Parameter APIKey
        APIKey to use for PushBullet. Defaults to using $MyAPIKey, which is specified at the top of this module
        .Outputs
        Nothing is sent to the pipeline. The result is that the locally cached files storing the valid Devices 
        and ContactEmails are updated.
        .Notes
        NAME:  Update-PushBulletSettings
        AUTHOR: Kevin Doblosky
        LASTEDIT: 2015-02-13
    #>
    param ($APIKey = $MyAPIKey)

    $emails = Get-PushbulletContacts -APIKey $APIKey | Select-Object email
    $emails | ConvertTo-Json | Set-Content (Join-Path $StorageLocation "PushBulletContacts.json")
     
    $devices = Get-PushbulletDevices | Select-Object nickname, iden 
    $devices | % { $_.nickname = $_.nickname -replace " ", "_" }
    $devices | ConvertTo-Json | Set-Content (Join-Path $StorageLocation "PushBulletDevices.json")
}


Function Get-PushbulletDevices
{
    <#
    .SYNOPSIS
        Get-PushbulletDevices will return the list of devices currently linked to your Pushbullet account.

    .DESCRIPTION
        This function requires an account at Pushbullet. Register at http://pushbullet.com and obtain your API Key from the account settings section.

        With this module you can retrieve a list of devices linked to your Pushbullet accounts, and then use the 'iden' value to send notifications to specific devices.
   
    .EXAMPLE
        Get-PushbulletDevices -APIKey "XXXXXX" | Select nickname,model,iden

        Get a table of names, models and iden numbers for all your devices.

    .EXAMPLE
        Get-PushbulletDevices -APIKey "XXXXXX" | Where {$_.nickname -eq "MyDevice"} | Select -ExpandProperty iden

        Return the iden value for a specific device called 'MyDevice'.
    #>
    param([string]$APIKey=$MyAPIKey)

    $Creds = New-Object System.Management.Automation.PSCredential ($APIKey, (ConvertTo-SecureString $APIKey -AsPlainText -Force))
    $devices = Invoke-WebRequest -Uri "https://api.pushbullet.com/v2/devices" -Credential $Creds -Method GET
    $json = $devices.content | ConvertFrom-Json
    return $json.devices
}

Function Get-PushbulletContacts
{
    <#
    .SYNOPSIS
        Get-PushbulletContacts will return your list of contacts.

    .DESCRIPTION
        This function requires an account at Pushbullet. Register at http://pushbullet.com and obtain your API Key from the account settings section.

        With this module you can see a list of your existing contacts, for use with Send-Pushbullet and the 'ContactEmail' parameter.
   
    .EXAMPLE
        Get-PushbulletContacts -APIKey "XXXXXX" | Select name,email

        Get a table of contact names and emails.
    #>
    param([string]$APIKey=$MyAPIKey)

    $Creds = New-Object System.Management.Automation.PSCredential ($APIKey, (ConvertTo-SecureString $APIKey -AsPlainText -Force))
    $devices = Invoke-WebRequest -Uri "https://api.pushbullet.com/v2/contacts" -Credential $Creds -Method GET
    $json = $devices.content | ConvertFrom-Json
    return $json.contacts
}

New-Alias -Name pushb -Value Send-PushbulletEx


Export-ModuleMember -Function Send-Pushbullet
Export-ModuleMember -Function Send-PushbulletEx
Export-ModuleMember -Function Update-PushBulletSettings
Export-ModuleMember -Function Get-PushbulletDevices
Export-ModuleMember -Function Get-PushbulletContacts
Export-ModuleMember -Alias pushb
