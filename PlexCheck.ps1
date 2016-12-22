<#
.SYNOPSIS
Pull a list of recently-added movies from Plex and send a listing via email

.DESCRIPTION
This script will send to a specified recipient a list of movies added to Plex in the past 7 days (or as specified).
This list will include information pulled dynamically from OMDBapi.com, the Open Movie Database.

.PARAMETERS
See param block for descriptions of available parameters

.EXAMPLE
PS C:\>PlexCheck.ps1 -Token xx11xx11xx1100xx0x01

.EXAMPLE
PS C:\>PlexCheck.ps1 -Token xx11xx11xx1100xx0x01 -Url 10.0.0.100 -Port 12345 -Days 14 -EmailTo test@test.com -Cred StoredCredential

.NOTES
To add credentials open up Control Panel>User Accounts>Credential Manager and click "Add a gereric credential". 
The "Internet or network address" field will be the Name required by the Cred param (default: "PlexCheck").

Requires StoredCredential.psm1 from https://gist.github.com/toburger/2947424, which in turn was adapted from
http://stackoverflow.com/questions/7162604/get-cached-credentials-in-powershell-from-windows-7-credential-manager

#>
param(
    # Required: specify your Plex Token
    #   To find your token, check here: https://support.plex.tv/hc/en-us/articles/204059436-Finding-your-account-token-X-Plex-Token
    [Parameter(Mandatory = $true)]
    [string]$Token,

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
    [string]$Cred = 'PlexCheck',

    # Specify the Library ID of any libraries you'd like to exclude
    [int[]]$ExcludeLib = 0,

    # Specify whether to prevent sending email if there are no additions
    [switch]$PreventSendingEmptyList,

    # Specify whether to omit the Plex Server version number from the email
    [switch]$OmitVersionNumber


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

$response = Invoke-WebRequest "$url`:$port/library/recentlyAdded/?X-Plex-Token=$Token" -Headers @{"accept"="application/json"}
$jsonlibrary = ConvertFrom-JSON $response.Content

# Grab those libraries!
$movies = $jsonLibrary.MediaContainer.Metadata |
    Where-Object {$_.type -eq 'movie' -AND $_.addedAt -gt (Get-Date (Get-Date).AddDays(-$days) -UFormat "%s")} |
    Select-Object * |
    Sort-Object addedAt

$tvShows = $jsonLibrary.MediaContainer.Metadata |
    Where-Object {$_.type -eq 'season' -AND $_.addedAt -gt (Get-Date (Get-Date).AddDays(-$days) -UFormat "%s")} |
    Group-Object parentTitle

# Initialize the counters and lists
$movieCount = 0
$movieList = "<h1>Movies:</h1><br/><br/>"
$movieList += "<table style=`"width:100%`">"
$tvCount = 0
$tvList = "<h1>TV Seasons:</h1><br/><br/>"
$tvList += "<table style=`"width:100%`">"

if ($($movies | Measure-Object).count -gt 0) {
    foreach ($movie in $movies) {
        # Make sure the movie's not in an excluded library
        if ($movie.librarySectionID -notin $ExcludeLib){
            $movieCount++

            # Retrieve movie info from the Open Movie Database
            $omdbURL = "omdbapi.com/?t=$($movie.title)&y=$($movie.year)&r=JSON"
            $omdbResponse = ConvertFrom-JSON (Invoke-WebRequest $omdbURL).content

            # If there was no result, try searching for the previous year (OMDB/The Movie Database quirkiness)
            if ($omdbResponse.Response -eq "False") {
                $omdbURL = "omdbapi.com/?t=$($movie.title)&y=$($($movie.year)-1)&r=JSON"
                $omdbResponse = ConvertFrom-JSON (Invoke-WebRequest $omdbURL).content
            }

            # If there was STILL no result, try searching for the *next* year
            if ($omdbResponse.Response -eq "False") {
                $omdbURL = "omdbapi.com/?t=$($movie.title)&y=$($($movie.year)+1)&r=JSON"
                $omdbResponse = ConvertFrom-JSON (Invoke-WebRequest $omdbURL).content
            }

            if ($omdbResponse.Response -eq "True") {
                if ($omdbResponse.Poster -eq "N/A") {
                    # If the poster was unavailable, substitute a Plex logo
                    $imgURL = $imgPlex
                    $imgHeight = "150"
                } else {
                    $imgURL = $omdbResponse.Poster
                    $imgHeight = "234"
                }
                $movieList += "<tr><td><img src=`"$imgURL`" height=$($imgHeight)px width=150px></td>"
                $movieList += "<td><li><a href=`"http://www.imdb.com/title/$($omdbResponse.imdbID)/`">$($movie.title)</a> ($($movie.year))</li>"
                $movieList += "<ul><li><i>Genre:</i> $($omdbResponse.Genre)</li>"
                $movieList += "<li><i>Rating:</i> $($omdbResponse.Rated)</li>"
                $movieList += "<li><i>Runtime:</i> $($omdbResponse.Runtime)</li>"
                $movieList += "<li><i>Director:</i> $($omdbResponse.Director)</li>"
                $movieList += "<li><i>Plot:</i> $($omdbResponse.Plot)</li>"
                $movieList += "<li><i>IMDB rating:</i> $($omdbResponse.imdbRating)/10</li>"
                $movieList += "<li><i>Added:</i> $(Get-Date $epoch.AddSeconds($movie.addedAt) -Format 'MMMM d')</li></ul></td>"
            }
            else {
                # If the movie couldn't be found in the DB even with the one-year buffer, fail gracefully
                $movieList += "<td><img src=`"$imgPlex`" height=150px width=150px></td><td><li>$($movie.title)</a> ($($movie.year)) - no additional information</li></td>"
            }
            $movieList += "</tr>"
        }
    }
    $movieList += "</table><br/><br/>"
}

if ($($tvShows | Measure-Object).Count -gt 0) {
    foreach ($show in $tvShows) {
        # Due to how shows are nested, gotta dig deep to get the librarySectionID
        if ($($show.group) -is [array]) {
            [int]$section = $($show.Group)[0].librarySectionID
        } else {
            [int]$section = $($show.Group).librarySectionID
        }

        # Make sure the media we're parsing isn't in an excluded library
        if (-not($ExcludeLib.Contains($section))){
             # Count it!
             $tvCount++

             # Retrieve show info from the Open Movie Database
             $omdbURL = "omdbapi.com/?t=$($show.name)&r=JSON"
             $omdbResponse = ConvertFrom-JSON (Invoke-WebRequest $omdbURL).content

             # Build the HTML
             if ($omdbResponse.Response -eq "True") {
                if ($omdbResponse.Poster -eq "N/A") {
                    # If the poster was unavailable, substitute a Plex logo
                    $imgURL = $imgPlex
                    $imgHeight = "150"
                } else {
                    $imgURL = $omdbResponse.Poster
                    $imgHeight = "234"
                }
                $tvList += "<tr><td><img src=`"$imgURL`" height=$($imgHeight)px width=150px></td>"
                $tvList += "<td><li><a href=`"http://www.imdb.com/title/$($omdbResponse.imdbID)/`">$($show.name)</a></li>"
                $tvList += "<ul><li><i>Genre:</i> $($omdbResponse.Genre)</li>"
                $tvList += "<li><i>Rating:</i> $($omdbResponse.Rated)</li>"
                $tvList += "<li><i>Plot:</i> $($omdbResponse.Plot)</li>"
                $tvList += "<li><i>Now in library:</i><br/></li><ul>"
                foreach ($season in ($show.Group | Sort-Object @{e={$_.index -as [int]}})){
                    if ($($season.leafCount) -gt 1) {
                        $plural = 's'
                    } else {
                        $plural = ''
                    }
                    $tvList += "<li>$($season.title) - $($season.leafCount) episode$($plural)</li>"
                }
                #$tvList += "<li><i>Added:</i> $(Get-Date $epoch.AddSeconds($movie.addedAt) -Format 'MMMM d')</li></ul></td>"
            }
            else {
                # If the series couldn't be found in the DB, fail gracefully
                $tvList += "<tr><td><img src=`"$imgPlex`" height=150px width=150px></td><td><li>$($show.name)</a></li>"
                            $tvList += "<td><li><a href=`"http://www.imdb.com/title/$($omdbResponse.imdbID)/`">$($show.name)</a></li>"
                $tvList += "<li><i>Season:</i><br/></li><ul>"
                foreach ($season in $show.Group){
                    $tvList += "<li>$($season.title) ($($season.leafCount) episode(s))</li>"
                }
            }
            $tvList += "</ul></ul></td></tr>"
        }
    }
    $tvList += "</table><br/>"
}



if (($movieCount -eq 0) -AND ($tvCount -eq 0)) {
    $body = "No movies or TV shows have been added to the Plex library in the past $days days. Sorry!"
} else {
    $body = "<h1>Hey there!</h1><br/>Here's the list of additions to my Plex library in the past $days days.<br/>"

    if ($movieCount -gt 0) {
        $body += $movieList
    }

    if ($tvCount -gt 0) {
        $body += $tvList
    }
    $body += "Enjoy!"
}

if (-not $OmitVersionNumber) {
    $body += "<br><br><br><br><p align = right><font size = 1 color = Gray>Plex Version: $((Invoke-RestMethod "$url`:$port/?X-Plex-Token=$Token" -Headers @{"accept"="application/json"}).mediaContainer.version)</p></font>"
}

$startDate = Get-Date (Get-Date).AddDays(-$days) -Format 'MMM d'
$endDate = Get-Date -Format 'MMM d'
    
$credentials = Get-StoredCredential -Name $cred
    
# If not otherwise specified, set the To address the same as the From
if ($EmailTo -eq 'default') {
    $EmailTo = $credentials.UserName
}
$subject = "Plex Additions from $startDate-$endDate"

if (-not($PreventSendingEmptyList -and (($movieCount+$tvCount) -eq 0))) {
    Send-MailMessage -From $($credentials.UserName) -to $EmailTo -SmtpServer $SMTPserver -Port $SMTPport -UseSsl -Credential $credentials -Subject $subject -Body $body -BodyAsHtml
}