# Add System.Web for urlencode
Add-Type -AssemblyName System.Web

[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")

#Module Imports
Import-Module ExchangeOnlineManagement

##############################################Exchange Online Utility Functions#############################################
# Function to Connect Exchange Online
function Connect-ExOnline {
    param ($UserName)

    Connect-ExchangeOnline -UserPrincipalName $UserName

    #Check for connectivity
    if ((Get-PSSession | Where-Object { $_.ConfigurationName -like "*Exchange*" }) -ne $null) {
        Write-Host `Successfully connected to Exchange Online -ForegroundColor Yellow
    }
    else {
        Write-Host `Not connected to Exchange Online -ForegroundColor Red
    }
}
###############################################################################################################################


##############################################Session Management Utility Functions#############################################
# Function to Get Exo Teams Session
function Get-Sessions {
    param (
        $ExoUserName,
        $TeamsUserName
    )

    try {
        Connect-ExOnline -UserName $ExoUserName
        Connect-MicrosoftTeams #-AccountId $TeamsUserName
    
    } catch {
        write-host "Error: $($_.Exception.Message)" -Foregroundcolor Red
    }
}

# Function To Remove existing sessions
function Remove-Sessions() {

    try {
        Get-PSSession | Remove-PSSession
        Write-Host All sessions in the current window has been removed. -ForegroundColor Yellow

    } catch {
        write-host "Error: $($_.Exception.Message)" -Foregroundcolor Red
    }
}
###############################################################################################################################



####################################################General Utility Functions##################################################

# Function To get system file content
function Get-SystemFileContent {
    param (
        $FileName
    )

    $Content = Get-Content $FileName -raw
    $Content = $Content -replace "(\\r)?\\n", [Environment]::NewLine
    $Content = ConvertFrom-StringData($Content)
    return $Content
}

# Function To Publish Error
function Publish-Error {
    param (
        $SC,
        $Member,
        $Message,
        $Log
    )

    $ErrorMessage = "Error($($SC)): $($Message) $($Member.Email)"
    Write-Host –ForegroundColor Red $ErrorMessage

    $Member | Select-Object GroupName, UserName, Email | Export-CSV $Log -NoTypeInformation -Append
}

# Function To Publish Success
function Publish-Success {
    param (
        $Member,
        $Message,
        $Log
    )

    $SuccessMessage = "Success: $($Message) $($Member.Email)"
    Write-Host –ForegroundColor Green $SuccessMessage

    if($Member.ChatId) {
        $Member | Select-Object GroupName, UserName, Email, UserId, ChatId, MsgId | Export-CSV $Log -NoTypeInformation -Append
    }
    else {
        $Member | Select-Object GroupName, UserName, Email | Export-CSV $Log -NoTypeInformation -Append
    }
}
###############################################################################################################################


####################################################MS Teams Utility Functions#################################################
# Function to Get Authentication Token
function Get-UserAccessToken {
    param (
        $ID,
        $Secret,
        $Scope,
        $UserName,
        $Password,
        $Tenant
    )
    
    # Authentication url
    $URL = "https://login.microsoftonline.com/" + $Tenant + "/oauth2/v2.0/token"

    # Create body
    $Body = @{
        client_id = $ID
	    client_secret = $Secret
	    scope = $Scope
	    grant_type = 'password'
        userName = $UserName
        password = $Password
    }

    # Splat the parameters for Invoke-Restmethod for cleaner code
    $Splat = @{
        ContentType = 'application/x-www-form-urlencoded'
        Method = 'POST'
        Body = $Body
        Uri = $URL
    }

    try {
        # Request the token!
        $Response = Invoke-RestMethod @splat

        # Create header
        $Header = @{
            Authorization = "$($Response.token_type) $($Response.access_token)"
        }
    } catch {
        $Header = $null
    }

    return $Header
}

# Function to Create New One2One Teams Chat
function New-One2OneChat {
    param (
        $Sender,
        $Recepient,
        $Header
    )

    # One on one chat creation url
    $URL = "https://graph.microsoft.com/beta/chats"

    # Create body
    $Body = '
    {
        "chatType": "oneOnOne",
        "members": [
            {
                "@odata.type": "#microsoft.graph.aadUserConversationMember",
                "roles": ["owner"],
                "user@odata.bind": "https://graph.microsoft.com/beta/users('+"'$Sender'"+')"
            },
            {
                "@odata.type": "#microsoft.graph.aadUserConversationMember",
                "roles": ["owner"],
                "user@odata.bind": "https://graph.microsoft.com/beta/users('+"'$Recepient'"+')"
            }
        ]
    }'

    try {
        # Invoke REST api to post chat
        $Response = Invoke-RestMethod -Uri $URL -Headers $Header -Body $Body -Method Post -ContentType "application/json"
    } catch {
        $Response = @{
            StatusCode = $_.Exception.Response.StatusCode.value__;
            StatusDescription = $_.Exception.Response.StatusDescription
        }
    }

    $Response
}

# Function to Send One2One Teams Chat
function Send-Chat {
    param (
        $ChatID,
        $Text,
        $Header
    )
    
    # Chat url
    $URL = "https://graph.microsoft.com/beta/chats/$($ChatID)/messages"

    # Create body
    $Body = "
    {
      'body': {
        'content': '$($Text)'
      }
    }"

    try {
        # Invoke REST api to post chat
        $Response = Invoke-RestMethod -Uri $URL -Headers $Header -Body $Body -Method Post -ContentType "application/json"
    } catch {       
        $Response = @{
            StatusCode = $_.Exception.Response.StatusCode.value__;
            StatusDescription = $_.Exception.Response.StatusDescription
        }
    }

    return $Response
}

# Function to Send Reminder
function Send-ChatActivityFeed {
    param (
        $Recepient,
        $ChatID,
        $Messageid,
        $Activity,
        $Preview,
        $Params,
        $Header
    )
    
    # Chat activity feed url
    $URL = "https://graph.microsoft.com/beta/chats/$($ChatID)/sendActivityNotification"

    # Create body
    $Body ="{
        'topic': {
            'source': 'entityUrl',
            'value': 'https://graph.microsoft.com/beta/chats/$($ChatID)/messages/$($Messageid)'
        },
        'activityType': '$($Activity)',
        'previewText': {
            'content': '$($Preview)'
        },
         'recipient': {
            '@odata.type': 'microsoft.graph.aadUserNotificationRecipient',
            'userId': '$($Recepient)'
        },
        'templateParameters': [$($Params)] 
    }"

    try {
        # Invoke REST api to post chat activity feed
        $Response = Invoke-RestMethod -Uri $URL -Headers $Header -Body $Body -Method Post -ContentType "application/json"
    } catch {
        $Response = @{
            StatusCode = $_.Exception.Response.StatusCode.value__;
            StatusDescription = $_.Exception.Response.StatusDescription
        }
    }

    return $Response
}

###############################################################################################################################


###########################################Sharepoint Online Utility Functions#################################################

# Function to Get Sharepoint Online Context
function Get-SpoContext {
    param (
        $SpoSiteUrl,
        $SpoUserName,
        $SpoPassword
    )
    
    try {
        $SecurePassword = $SpoPassword | ConvertTo-SecureString -AsPlainText -Force
        $Context = New-Object Microsoft.SharePoint.Client.ClientContext($SpoSiteUrl)
        $Context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($SpoUserName, $SecurePassword)
        return $Context

    } catch {
        write-host "Error: $($_.Exception.Message)" -Foregroundcolor Red
    }
}

# Function to Download Sharepoint Online File
function Get-SpoFile {
    param (
        $SpoContext,
        $SourceFileURL,
        $TargetFilePath
    )
       
    try {

        $SourceFileName = Split-path $SourceFileURL -leaf

        $SourceFile = $SpoContext.web.GetFileByUrl($SourceFileURL)
        $SpoContext.Load($SourceFile)
        $SpoContext.ExecuteQuery()

        [Microsoft.SharePoint.Client.FileInformation] $SourceFileInfo = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($SpoContext,$SourceFile.ServerRelativeUrl);
        [System.IO.FileStream] $TargetFileStream = [System.IO.File]::Open($TargetFilePath,[System.IO.FileMode]::Create);   
        $SourceFileInfo.Stream.CopyTo($TargetFileStream);
        $TargetFileStream.Close();

    } catch {
        write-host "Error: $($_.Exception.Message)" -Foregroundcolor Red
    }
}

# Function to Get Sharepoint Online File Content
function Get-SpoFileContent {
    param (
        $SpoContext,
        $SourceFileURL
    )
       
    try {

        $SourceFileName = Split-path $SourceFileURL -leaf

        $SourceFile = $SpoContext.web.GetFileByUrl($SourceFileURL)
        $SpoContext.Load($SourceFile)
        $SpoContext.ExecuteQuery()

        $TargetFilePath = "$(Get-Location)\$($SourceFileName)"

        [Microsoft.SharePoint.Client.FileInformation] $SourceFileInfo = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($SpoContext,$SourceFile.ServerRelativeUrl);
        [System.IO.FileStream] $TargetFileStream = [System.IO.File]::Open($TargetFilePath,[System.IO.FileMode]::Create);   
        $SourceFileInfo.Stream.CopyTo($TargetFileStream);
        $TargetFileStream.Close();

        $Content = Get-Content $SourceFileName -raw
        $Content = $Content -replace "(\\r)?\\n", [Environment]::NewLine
        $Content = ConvertFrom-StringData($Content)

        If(Test-Path $TargetFilePath) { Remove-Item $TargetFilePath}

        return $Content

    } catch {
        write-host "Error: $($_.Exception.Message)" -Foregroundcolor Red
    }
}

# Function to save single file to Sharepoint Online document library
function Send-SingleFileToSpo {
    param (
        $SpoContext,
        $DocLibraryName,
        $DestinationPath,
        $SourceFilePath
    )

    try {
        #Get the Library
        $DocLibrary =  $SpoContext.Web.Lists.GetByTitle($DocLibraryName)
 
        #Get file from disk
        $FileStream = ([System.IO.FileInfo] (Get-Item $SourceFilePath)).OpenRead()

        #Get source file Name from source file path
        $SourceFileName = Split-path $SourceFilePath -leaf
   
        #Upload single file to sharepoint online document library
        $FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
        $FileCreationInfo.Overwrite = $true
        $FileCreationInfo.ContentStream = $FileStream
        $FileCreationInfo.URL = $SourceFileName
        $FileUploaded = $DocLibrary.RootFolder.Files.Add($FileCreationInfo)
  
        $SpoContext.Load($FileUploaded)
        $SpoContext.ExecuteQuery()
 
        #Close source file stream
        $FileStream.Close()

    } catch {
        write-host "Error: $($_.Exception.Message)" -Foregroundcolor Red
    }
}

# Function to Get List Items By List Title
function Get-SpoListItems {
   param (
        [Microsoft.SharePoint.Client.ClientContext]$SpoContext,
        [String]$ListTitle
    )

    try {
        $List = $SpoContext.Web.Lists.GetByTitle($ListTitle)
        $Query = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()
        $ListItems = $List.GetItems($Query)
        $SpoContext.Load($ListItems)
        $SpoContext.ExecuteQuery()
        return $ListItems 

    } catch {
        write-host "Error: $($_.Exception.Message)" -Foregroundcolor Red
    }
}

###############################################################################################################################


