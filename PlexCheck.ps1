<#
.SYNOPSIS
Pull a list of recently-added movies from Plex and send a listing via email

.DESCRIPTION
This script will send to a specified recipient a list of movies added to Plex in the past 7 days (or as specified).
This list will include information pulled dynamically from OMDBapi.com, the Open Movie Database.

.PARAMETERS
See param block for descriptions of available parameters

.EXAMPLE
PS C:\>PlexCheck.ps1

.EXAMPLE
PS C:\>PlexCheck.ps1 -Url 10.0.0.100 -Port 12345 -Days 14 -EmailTo test@test.com -Cred StoredCredential

.NOTES
To add credentials open up Control Panel>User Accounts>Credential Manager and click "Add a gereric credential". 
The "Internet or network address" field will be the Name required by the Cred param (default: "PlexCheck").

Requires StoredCredential.psm1 from https://gist.github.com/toburger/2947424, which in turn was adapted from
http://stackoverflow.com/questions/7162604/get-cached-credentials-in-powershell-from-windows-7-credential-manager

#>
param(
    # Optionally specify IP of the server we want to connect to
    [string]$Url = 'http://127.0.0.1',

    # Optionally define a custom port
    [int]$Port = '32400',

    # Optionally specify a number of days back to report
    [int]$Days = 7,

    # Optionally define the address to send the report to
    # If not otherwise specified, send to the From address
    [string]$EmailTo = 'default',

    # Specify the SMTP server address (if not gmail)
    # Assumes SSL, because security!
    [string]$SMTPserver = 'smtp.gmail.com',

    # Specify the SMTP server's SSL port
    [int]$SMTPport = '587',

    # Specify the name used for the Credential Manager entry
    [string]$Cred = 'PlexCheck'


)

#region Associated Files
if (-not (Get-Module Get-CredentialFromWindowsCredentialManager)) {
    Try {
        Import-Module Get-CredentialFromWindowsCredentialManager.psm1 -ErrorAction Stop
    } Catch {
        Write-Host "Failed to load Get-CredentialFromWindowsCredentialManager.psm1. Aborting."
        Exit
    }
}
#endregion

#region Declarations
$epoch = Get-Date '1/1/1970'
$imgPlex = "http://i.imgur.com/RyX9y3A.jpg"
#endregion



Invoke-WebRequest "$url`:$port/library/recentlyAdded/" -OutFile "$PSScriptRoot\library.xml"
[xml]$library = Get-Content -Path "$PSScriptRoot\library.xml"

$movies = $library.SelectNodes("/MediaContainer/Video") |
    Where-Object {$_.addedAt -gt (Get-Date (Get-Date).AddDays(-$days) -UFormat "%s")} |
    Select-Object * |
    Sort-Object addedAt

$body = "Hey there!<br/><br/>The movies listed below were added to the <a href=`"http://app.plex.tv/web/app`">Plex library</a> in the past $days days.<br/><br/>"
$body += '<table style="width:100%">'

foreach ($movie in $movies){
    $omdbURL = "omdbapi.com/?t=$($movie.title)&y=$($movie.year)&r=JSON"
    $body += "<tr>"
    $omdbResponse = ConvertFrom-JSON (Invoke-WebRequest $omdbURL).content
    if ($omdbResponse.Response -eq "True") {
        if ($omdbResponse.Poster -eq "N/A") {
            # If the poster was unavailable, substitute a Plex logo
            $imgURL = $imgPlex
            $imgHeight = "150"
        } else {
            $imgURL = $omdbResponse.Poster
            $imgHeight = "234"
        }
        $body += "<td><img src=`"$imgURL`" height=$($imgHeight)px width=150px></td>"
        $body += "<td><li><a href=`"http://www.imdb.com/title/$($omdbResponse.imdbID)/`">$($movie.title)</a> ($($movie.year))</li>"
        $body += "<ul><li><i>Genre:</i> $($omdbResponse.Genre)</li>"
        $body += "<li><i>Rating:</i> $($omdbResponse.Rated)</li>"
        $body += "<li><i>Runtime:</i> $($omdbResponse.Runtime)</li>"
        $body += "<li><i>Director:</i> $($omdbResponse.Director)</li>"
        $body += "<li><i>Plot:</i> $($omdbResponse.Plot)</li>"
        $body += "<li><i>IMDB rating:</i> $($omdbResponse.imdbRating)/10</li>"
        $body += "<li><i>Added:</i> $(Get-Date $epoch.AddSeconds($movie.addedAt) -Format 'MMMM d')</li></ul></td>"
    }
    else {
        # If the movie couldn't be found in the DB, fail gracefull
        $body += "<td><img src=`"$imgPlex`" height=150px width=150px></td><td><li>$($movie.title)</a> ($($movie.year)) - no additional information</li></td>"
    }
    $body += "</tr>"
    
}
$body += "</table><br/>Enjoy!"

$startDate = Get-Date (Get-Date).AddDays(-$days) -Format 'MMM d'
$endDate = Get-Date -Format 'MMM d'

$credentials = Get-StoredCredential -Name $cred

# If not otherwise specified, set the To address the same as the From
if ($EmailTo -eq 'default') {
    $EmailTo = $credentials.UserName
}
$subject = "Plex Additions from $startDate-$endDate"

Send-MailMessage -From $($credentials.UserName) -to $EmailTo -SmtpServer $SMTPserver -Port $SMTPport -UseSsl -Credential $credentials -Subject $subject -Body $body -BodyAsHtml
