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
PS C:\>PlexCheck.ps1 -Url 10.0.0.100 -Port 12345 -Days 14 -EmailTo test@test.com -ExcludeLib 11 -PreventSendingEmptyList -OmitVersionNumber

.NOTES
    Requires CredentialManager module.

    > Install-Module CredentialManager

    > New-StoredCredential -Target plexToken -UserName plex -Password [Plex token] -Type Generic -Persist LocalMachine
    > New-StoredCredential -Target tmdb.org -UserName tmdb -Password [TMDB token] -Type Generic -Persist LocalMachine
    > New-StoredCredential -Target PlexCheck -UserName [Email address] -Password [Email password] -Type Generic -Persist LocalMachine

    To find your Plex token, check here: https://support.plex.tv/hc/en-us/articles/204059436-Finding-your-account-token-X-Plex-Token

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

    # Specify the Library ID of any libraries you'd like to exclude
    [int[]]$ExcludeLib = @(),

    # Specify whether to prevent sending email if there are no additions
    [switch]$PreventSendingEmptyList,

    # Specify whether to omit the Plex Server version number from the email
    [switch]$OmitVersionNumber
)

#region import credentials
if (-not (Get-Module CredentialManager)) {
    Try {
        Import-Module CredentialManager -ErrorAction Stop
    } Catch {
        Write-Host "Failed to load CredentialManager. Aborting."
        Throw $_.Exception
    }
}

Try {
    $tmdbToken = (Get-StoredCredential -Target tmdb.org -ErrorAction Stop).GetNetworkCredential().Password
    if (-not $tmdbToken) {
        Throw "No TMDB token found."
    }
    Write-Verbose "Retrieved TMDB token."
}
Catch {
    Write-Error "Failed to retrieve TMDB token."
    Throw $_.Exception
}

Try {
    $plexToken = (Get-StoredCredential -Target PlexToken -ErrorAction Stop).GetNetworkCredential().Password
    if (-not $plexToken) {
        Throw "No Plex token found."
    }
    Write-Verbose "Retrieved Plex token."
}
Catch {
    Write-Error "Failed to retrieve Plex token."
    Throw $_.Exception
}

Try {
    $emailCreds = Get-StoredCredential -Target PlexCheck -ErrorAction Stop
    if (-not $emailCreds) {
        Throw "No email credentials found."
    }
    Write-Verbose "Retrieved email credentials."
}
Catch {
    Throw $_.Exception
}

#endregion

#region Declarations
$epoch = Get-Date '1/1/1970'
$startDate = Get-Date (Get-Date).AddDays(-$days) -UFormat "%s"
$imgPlex = "http://i.imgur.com/RyX9y3A.jpg"
$searchURL = "https://api.themoviedb.org/3/find"
$imdbIDformat = [Regex]::new('tt\d{7,8}')
$tvdbIDformat = [Regex]::new('[1-9]\d*')
#endregion

$response = Invoke-RestMethod "$url`:$port/library/recentlyAdded/?X-Plex-Token=$plexToken" -Headers @{"accept"="application/json"} | Select-Object -ExpandProperty MediaContainer | Select-Object -ExpandProperty Metadata

# Grab those libraries!
$movies = $response |
    Where-Object {$_.type -eq 'movie' -AND $_.addedAt -gt $startDate} |
    Select-Object * |
    Sort-Object addedAt

$tvShows = $response |
    Where-Object {$_.type -eq 'season' -AND $_.addedAt -gt $startDate} |
    Group-Object parentTitle

# Initialize the counters and lists
$movieCount = 0
$movieList = "<h1>Movies:</h1><br/><br/>"
$movieList += "<table style=`"width:100%`">"

if ($($movies | Measure-Object).count -gt 0) {
    foreach ($movie in $movies | Where-Object {$_.librarySectionID -notin $ExcludeLib}) {
        # TMDB rate-limits. Currently 40 requests every 10 seconds, so sleep 10 seconds every 15 movies to be safe (2 calls per movie)
        # https://developers.themoviedb.org/3/getting-started/request-rate-limiting
        if ($movieCount -gt 1 -AND $movieCount%15 -eq 0) {
            Write-Verbose "On movie $movieCount, waiting 10 seconds..."
            Start-Sleep -Seconds 10
        }

        $movieCount++

        Write-Verbose "Looking up $($movie.title) ($movieCount / $($movies | Measure-Object | Select-Object -ExpandProperty Count))."

        # Retrieve movie info from The Movie Database
        $simpleResponse = (Invoke-RestMethod "$searchURL/$($imdbIDformat.Matches($movie.guid).value)?api_key=$tmdbToken&language=en-US&external_source=imdb_id").movie_results

        # Assuming we have a valid response, pull detailed info on the movie
        if ($simpleResponse.id) {
            $detailedResponse = (Invoke-RestMethod "https://api.themoviedb.org/3/movie/$($simpleResponse.id)?api_key=$tmdbToken&language=en-US")
        }

        if ($detailedResponse.Title) {
            if ($detailedResponse.poster_path) {
                $movieList += "<tr><td><img src=`"https://image.tmdb.org/t/p/w154$($detailedResponse.poster_path)`"</td>"
            } else {
                # If the poster was unavailable, substitute a Plex logo
                $movieList += "<tr><td><img src=`"$imgPlex`" height=154px width=154px></td>"
            }
            $movieList += "<td><li><a href=`"http://www.imdb.com/title/$($detailedResponse.imdb_ID)/`">$($detailedResponse.title)</a> ($(($detailedResponse.release_date).split('-')[0]))</li>"
            if($detailedResponse.Genres.count -gt 1) {
                $movieList += "<li><i>Genres:</i> $($detailedResponse.Genres.name -join ', ')</li>"
            } elseif ($detailedResponse.Genres.count -eq 1) {
                $movieList += "<li><i>Genre:</i> $($detailedResponse.Genres.name)</i>"
            }
            $movieList += "<li><i>Rating:</i> $($movie.contentRating)</li>"
            $movieList += "<li><i>Runtime:</i> $($detailedResponse.runtime) minutes</li>"
            if($movie.role.count -gt 1) {
                $movieList += "<li><i>Stars:</i> $($movie.role.tag -join ', ')</li>"
            } elseif ($movie.role.count -eq 1) {
                $movieList += "<li><i>Star:</i> $($movie.role.tag)</i>"
            }
            $movieList += "<li><i>Plot:</i> $($movie.summary)</li>"
            $movieList += "<li><i>IMDB rating:</i> $(($detailedResponse.vote_average).toString("#.#"))/10</li>"
            $movieList += "<li><i>Added:</i> $(Get-Date $epoch.AddSeconds($movie.addedAt) -Format 'MMMM d')</li></ul></td>"
        }
        else {
            # If the movie couldn't be found in the DB even with the one-year buffer, fail gracefully
            $movieList += "<tr><td><img src=`"$imgPlex`" height=150px width=150px></td><td><li>$($movie.title)</a> ($($movie.year)) - no additional information</li></td>"
        }
        $movieList += "</tr>"

        Clear-Variable simpleResponse, detailedResponse
    }
    $movieList += "</table><br/><br/>"
}

$tvCount = 0
$tvList = "<h1>TV Seasons:</h1><br/><br/>"
$tvList += "<table style=`"width:100%`">"

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
            if ($movieCount -gt 1 -AND $movieCount%15 -eq 0) {
                Write-Verbose "On movie $movieCount, waiting 10 seconds..."
                Start-Sleep -Seconds 10
            }

            # Count it!
            $tvCount++

            $tvdbID = ($tvdbIDformat.matches($show.group.guid).value)[0]

            Write-Verbose "Looking up $($show.name) ($tvCount / $($tvShows | Measure-Object | Select-Object -ExpandProperty Count))."

            # Retrieve movie info from The Movie Database
            $simpleResponse = (Invoke-RestMethod "$searchURL/$tvdbID`?api_key=$tmdbToken&language=en-US&external_source=tvdb_id").tv_results

            # Assuming we have a valid response, pull detailed info on the movie
            if ($simpleResponse.id) {
                $detailedResponse = (Invoke-RestMethod "https://api.themoviedb.org/3/tv/$($simpleResponse.id)?api_key=$tmdbToken&language=en-US")
                $contentRating = (Invoke-RestMethod "https://api.themoviedb.org/3/tv/$($simpleResponse.id)/content_ratings?api_key=$tmdbToken&language=en-US").Results | Where-Object {$_.iso_3166_1 -eq 'US'} | Select-Object -ExpandProperty Rating
                $imdbID = (Invoke-RestMethod "https://api.themoviedb.org/3/tv/$($simpleResponse.id)/external_ids?api_key=$tmdbToken&language=en-US").imdb_id
            }

            if ($detailedResponse.id) {
                if ($detailedResponse.poster_path) {
                    $tvList += "<tr><td><img src=`"https://image.tmdb.org/t/p/w154$($detailedResponse.poster_path)`"</td>"
                } else {
                    # If the poster was unavailable, substitute a Plex logo
                    $tvList += "<tr><td><img src=`"$imgPlex`" height=154px width=154px></td>"
                }
                $tvList += "<td><li><a href=`"http://www.imdb.com/title/$imdbID/`">$($show.name)</a></li>"
                if ($detailedResponse.genres) {$tvList += "<ul><li><i>Genre:</i> $($detailedResponse.genres.name -join ", ")</li>"}
                if ($contentRating) {$tvList += "<li><i>Rating:</i> $contentRating</li>"}
                $tvList += "<li><i>Plot:</i> $($detailedResponse.overview)</li>"
                $tvList += "<li><i>Now available:</i><br/></li><ul>"
                foreach ($season in ($show.Group | Sort-Object @{e={$_.index -as [int]}})){
                    $tvList += "<li>$($season.title): $($season.leafCount) episode$(if ($season.leafCount -gt 1){"s"})</li>"
                }
            }
            else {
                # If the series couldn't be found in the DB, fail gracefully
                $tvList += "<tr><td><img src=`"$imgPlex`" height=150px width=150px></td><td><li>$($show.name)</a></li>"
                $tvList += "<td><li><a href=`"http://www.imdb.com/title/$($omdbResponse.imdbID)/`">$($show.name)</a></li>"
                foreach ($season in $show.Group){
                    $tvList += "<li>$($season.title) ($($season.leafCount) episode$(if ($season.leafCount -gt 1){"s"})</li>"
                }
            }
            $movieList += "</tr>"

            Clear-Variable simpleResponse, detailedResponse, contentRating, imdbID, season

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
    $body += "<br><br><br><br><p align = right><font size = 1 color = Gray>Plex Version: $((Invoke-RestMethod "$url`:$port/?X-Plex-Token=$plexToken" -Headers @{"accept"="application/json"}).mediaContainer.version). Posters/metadata from TMDb.</p></font>"
}

$startDate = Get-Date (Get-Date).AddDays(-$days) -Format 'MMM d'
$endDate = Get-Date -Format 'MMM d'

    
# If not otherwise specified, set the To address the same as the From
if ($EmailTo -eq 'default') {
    $EmailTo = $emailCreds.UserName
}
$subject = "Plex Additions from $startDate-$endDate"

if (-not($PreventSendingEmptyList -and (($movieCount+$tvCount) -eq 0))) {
    Send-MailMessage -From $($emailCreds.UserName) -to $EmailTo -SmtpServer $SMTPserver -Port $SMTPport -UseSsl -Credential $emailCreds -Subject $subject -Body $body -BodyAsHtml
}