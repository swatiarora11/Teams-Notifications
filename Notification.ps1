Set-ExecutionPolicy -ExecutionPolicy Bypass -SCope CurrentUser
Clear-Host

Import-Module -Name "./NotificationUtil.psm1" -Verbose

# Function to Clean Up Post Script Run
function CleanUp {
    param (
        $SpoContext
    )

    ##Dispose Sharepoint Online context
    $SpoContext.Dispose()
    
    ##Remove Current Sessions
    Remove-Sessions
}

# Function to Get User Choice
function Get-UserChoice {
    Write-Host ""
    Write-Host This script is for communication with O365 group members within your company tenant. -ForegroundColor Yellow
    Write-Host ""
    Write-Host ========================================================================================================================
    Write-Host "                                                  SCRIPT CONFIGURATION"
    Write-Host ========================================================================================================================
    Write-Host ""
    Write-Host "Exchange Online User   :" $Configuration.'exo.user.name'
    Write-Host "Notifying User         :" $Configuration.'teams.user.name'
    Write-Host ========================================================================================================================
    Write-Host ""
    Write-Host "Press Y: to proceed N: to cancel and quit" -ForegroundColor Yellow

    $Option=Read-Host "Enter your choice?"
    Switch ($Option)
    {
	    Y {$Choice="Proceed"} 
	    Default {$Choice="Quit"}
    }

    return $Choice
}

#Function to Send Notification to Group Memebers
function Send-NotificationToGroupMembers {
   param (
        $SpoContext,
        $GroupIds,
        $Notifier,
        $Notification,
        $Config,
        $LogID
    )

    $CurrentFolder = Get-Location

    $SLog = "$($CurrentFolder)\Log-$($LogID)-NS.csv"
    $FLog = "$($CurrentFolder)\Log-$($LogID)-NF.csv"

    If(Test-Path $SLog) { Remove-Item $SLog}
    If(Test-Path $FLog) { Remove-Item $FLog}

    #Do for each EXO group as given in the configuration
    ForEach ($GID in $GroupIds) 
    {
        $GID = $GID.Trim()
        $Group = Get-UnifiedGroup –Identity $GID
        Write-Host "Group Name   :" $Group.DisplayName 
        Write-Host "Group Members:" $Group.GroupMemberCount 
        Write-Host "Group Type   :" $Group.GroupType

        $GroupMembers = Get-UnifiedGroupLinks –Identity $GID –LinkType Members | Select-Object @{Name="GroupName";Expression={$Group.DisplayName}},`
        @{Name="UserName";Expression={$_.DisplayName}}, @{Name="Email";Expression={$_.PrimarySmtpAddress}}, @{Name="UserId";Expression={$_.ExternalDirectoryObjectId}}

        #Do for each member of the group
        ForEach ($Member in $GroupMembers) 
        {
            $User = $Member.UserId
            if($Notifier -eq $User) { continue }

            #Create teams one on one chat
            $NewOne2OneChatResponse = New-One2OneChat -Header $Header -Sender $Notifier -Recepient $User
            if($NewOne2OneChatResponse['StatusCode']) {
                Publish-Error -SC $NewOne2OneChatResponse['StatusCode'] -Member $Member -Message "Unable to create one2one chat to user" -Log $FLog
                continue 
            }
  
            if ($NewOne2OneChatResponse.id)
            {
                #Send teams chat
                $SendChatResponse = Send-Chat -Header $Header -ChatID $NewOne2OneChatResponse.id -Text $Notification
               
                if($SendChatResponse['StatusCode']) {
                    Publish-Error -SC $SendChatResponse['StatusCode'] -Member $Member -Message "Unable to send one2one chat to user" -Log $FLog
                    continue 
                }
            }

            $Table = @{}
            $Table.Add('GroupName', $member.GroupName)
            $Table.Add('UserName', $member.UserName)
            $Table.Add('Email', $member.Email)
            $Table.Add('UserId', $User)
            $Table.Add('ChatId', $NewOne2OneChatResponse.id)
            $Table.Add('MsgId', $SendChatResponse.id)

            $GroupMember = New-Object -TypeName PSObject -Property $Table

            Publish-Success -Member $GroupMember -Message "Notification sent to user" -Log $SLog
        }

        Write-Host ""
    }

    If(Test-Path $SLog) { 
        Send-SingleFileToSpo -SpoContext $SpoContext -DocLibraryName "TeamsNotification" -DestinationPath "Logs" -SourceFilePath $SLOG
        Remove-Item $SLog
    }

    If(Test-Path $FLog) { 
        Send-SingleFileToSpo -SpoContext $SpoContext -DocLibraryName "TeamsNotification" -DestinationPath "Logs" -SourceFilePath $FLOG
        Remove-Item $FLog
    }
}

#Function to Send Reminder to Group Memebers
function Send-Reminder {
   param (
        $SpoContext,
        $Notifier,
        $Config,
        $LogID,
        $PreviewText
    )

    $CurrentFolder = Get-Location

    $SLog = "$($CurrentFolder)\Log-$($LogID)-RS.csv"
    $FLog = "$($CurrentFolder)\Log-$($LogID)-RF.csv"

    If(Test-Path $SLog) { Remove-Item $SLog}
    If(Test-Path $FLog) { Remove-Item $FLog}

    $LogFileName = "Log-$($LogID)-NS.csv"
    $LogFilePath = "$($CurrentFolder)\$($LogFileName)"

    Get-SpoFile -SpoContext $SpoContext -SourceFileURL "$($SpoContext.Url)/TeamsNotification/$($LogFileName)" -TargetFilePath $LogFilePath

    $NotificationObjects = Import-Csv "$(Get-Location)\$($LogFileName)" | ForEach-Object {
        [PSCustomObject]@{
            'GroupName' = $_.GroupName
            'UserName' = $_.UserName
            'Email' = $_.Email
            'UserId' = $_.UserId
            'ChatId' = $_.ChatId
            'MsgId' = $_.MsgId
        }
    }

    If(Test-Path $LogFilePath) { Remove-Item $LogFilePath}

    #Do for each chat object 
    ForEach ($Notification in $NotificationObjects) 
    {
        $User = $Notification.UserId
        if($Notifier -eq $User) { continue }

        if ($Notification.ChatId)
        {
            #Send teams chat activity feed
            $SendChatActivityResponse = Send-ChatActivityFeed -Header $Header -Recepient $User -ChatID $Notification.ChatId `
            -Messageid $Notification.MsgId -Preview $PreviewText -Activity "approvalRequired" `
            -Params "{'name': 'deploymentId', 'value': '12345'}"

            if($SendChatActivityResponse['StatusCode']) { 
                Publish-Error -SC $SendChatActivityResponse['StatusCode'] -Member $Notification -Message "Unable to send reminder to user" -Log $FLog
                continue 
            }
        }

        Publish-Success -Member $Notification -Message "Reminder sent to user" -Log $SLog
    }

    Write-Host ""
    If(Test-Path $SLog) { 
        Send-SingleFileToSpo -SpoContext $SpoContext -DocLibraryName "TeamsNotification" -DestinationPath "Logs" -SourceFilePath $SLOG
        Remove-Item $SLog
    }

    If(Test-Path $FLog) { 
        Send-SingleFileToSpo -SpoContext $SpoContext -DocLibraryName "TeamsNotification" -DestinationPath "Logs" -SourceFilePath $FLOG
        Remove-Item $FLog
    }
}

#Get Configuration from system file
$Configuration = Get-SystemFileContent -FileName ".\app.properties"

##Get Sharepoint Online Context
$SpoContext = Get-SpoContext -SpoSiteUrl $Configuration.'spo.site.url' `
                    -SpoUserName $Configuration.'spo.user.name' `
                    -SpoPassword $Configuration.'spo.user.password'

#Get Configuration from Sharepoint Online
#$RemoteConfiguration = Get-SpoFileContent -SpoContext $SpoContext -SourceFileURL "$($SpoContext.Url)/TeamsNotification/app2.properties"

#Get Exchange Online and Teams Sessions
Get-Sessions -ExoUserName $Configuration.'exo.user.name' -TeamsUserName $Configuration.'teams.user.name'

#Get O365 user access token
$Header = Get-UserAccessToken -id $Configuration.'app.id' -secret $Configuration.'app.secret' -Scope $Configuration.'app.scope' `
-UserName $Configuration.'teams.user.name' -Password $Configuration.'teams.user.password' -Tenant $Configuration.'tenant.name'
if($Header -eq $null) {
    $ErrorMessage = "Error: Unable to obtain user access token"
    Write-Host –ForegroundColor Red $ErrorMessage
    return
}

$Choice=Get-UserChoice
if($Choice -eq 'Proceed') {
    
    Write-Host "Proceeding..." -ForegroundColor Yellow
    $NotificationList = Get-SpoListItems -SpoContext $SpoContext -ListTitle "Notification"

    foreach($item in $NotificationList)
    {
        $Notification = @{
            ID = ($item.FieldValues.ID)
            Title = ($item.FieldValues.Title)
            Notification = ($item.FieldValues.Notification)
            ExoGroupIds = ($item.FieldValues.ExoGroupIds -split ',')
            Send = ($item.FieldValues.Send)
            Remind = ($item.FieldValues.Remind)
            PreviewText = ($item.FieldValues.PreviewText)
        }

        if($Notification.Send) {

            Write-Host "Executing Script to Send Notification Title $($Notification.Title)" -ForegroundColor Yellow

            Send-NotificationToGroupMembers -SpoContext $SpoContext -Config $Configuration -LogID $Notification.ID `
            -GroupIds $Notification.ExoGroupIds -Notifier $Configuration.'teams.user.id' -Notification $Notification.Notification
        } 
        elseif($Notification.Remind) {
    
            Write-Host "Executing Script to Send Reminder for Notification Title $($Notification.Title)" -ForegroundColor Yellow

            Send-Reminder -SpoContext $SpoContext -Config $Configuration -LogID $Notification.ID `
            -Notifier $Configuration.'teams.user.id' -PreviewText $Notification.PreviewText
        }
    }
}
elseif($Choice -eq 'Quit') {

    Write-Host "Quitting..." -ForegroundColor Yellow
}

CleanUp -SpoContext $SpoContext


