<#
.SYNOPSIS
Pull a list of recently-added movies from Plex and send a listing via email

.DESCRIPTION
This script will send to a specified recipient a list of movies added to Plex in the past 7 days (or as specified).
This list will include information pulled dynamically from OMDBapi.com, the Open Movie Database.

.PARAMETERS
See param block for descriptions of available parameters

.EXAMPLE
PS C:\>Send-PlexLibraryUpdate.ps1

.EXAMPLE
PS C:\>Send-PlexLibraryUpdate.ps1 -Url 10.0.0.100 -Port 12345 -Days 14 -EmailTo test@test.com -ExcludeLib 11 -PreventSendingEmptyList -OmitVersionNumber

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
    [string] $PlexUrl = 'http://127.0.0.1',

    # Optionally define a custom port
    [int] $PlexPort = '32400',

    # Optionally specify a number of days back to report
    [int] $Days = 7,

    # Optionally define the address to send the report to
    # If not otherwise specified, send to the From address
    [string] $EmailTo = 'default',

    # Specify the SMTP server address (if not gmail)
    # Assumes SSL, because security!
    [string] $SMTPserver = 'smtp.gmail.com',

    # Specify the SMTP server's SSL port
    [int] $SMTPport = '587',

    # Specify the Library Name of any libraries you'd like to exclude
    [string[]] $ExcludeLib = @(),

    # Specify whether to prevent sending email if there are no additions
    [switch] $PreventSendingEmptyList,

    # Specify whether to omit the Plex Server version number from the email
    [switch] $OmitVersionNumber,

    # Specify whether to pick a random collection to feature
    [switch] $FeatureRandomCollection,

    # Specify whether to upload all movie posters to Imgur rather than use TVDB posters
    [Parameter (ParameterSetName = 'imgur')]
    [switch] $UploadPostersToImgur,

    # Specify the location of the file to store Imgur links in for movies/collections
    [Parameter (Mandatory = $true,
        ParameterSetName = 'imgur')]
    [ValidateScript ({
        if(-Not ($_ | Test-Path) ){
            throw "File or folder does not exist"
        }
        if(-Not ($_ | Test-Path -PathType Leaf) ){
            throw "The Path argument must be a file. Folder paths are not allowed."
        }
        if($_ -notmatch "(\.csv)"){
            throw "The file specified in the path argument must be either of type msi or exe"
        }
        return $true
    })]
    [System.IO.FileInfo] $ImgurFilePath
)

#region Invoke-ImgurUpload
function Invoke-ImgurUpload {
    param (
        [Parameter (Mandatory = $true)]
        [string] $ClientID,

        [Parameter (Mandatory = $true)]
        [string] $ImageBase64,

        [Parameter (Mandatory = $true)]
        [ValidateScript ({
            if(-Not ($_ | Test-Path) ){
                throw "File or folder does not exist"
            }
            if(-Not ($_ | Test-Path -PathType Leaf) ){
                throw "The Path argument must be a file. Folder paths are not allowed."
            }
            if($_ -notmatch "(\.csv)"){
                throw "The file specified in the path argument must be either of type msi or exe"
            }
            return $true
        })]
        [System.IO.FileInfo] $FilePath,

        [Parameter (Mandatory = $true)]
        [string] $PlexPath
    )

    Try {
        $csv = Import-CSV -Path $FilePath -ErrorAction Stop
    }
    Catch {
        Throw $_.Exception
    }

    if ($csv.PlexPath -contains $PlexPath) {
        Write-Verbose "Item already exists in $FilePath. Returning existing Imgur record."
        Return ($csv | Where-Object {$_.PlexPath -eq $PlexPath})
    }
    else {
        Write-Verbose "Item does not exist in $FilePath. Uploading to Imgur."
        $imgurHeader = @{Authorization = "Client-ID $ClientID"; "Content-Type" = "application/json"}

        Try {
            $imgurResponse = Invoke-RestMethod -Method Post -Uri "https://api.imgur.com/3/upload" -Headers $imgurHeader -Body (@{image = "$ImageBase64"; type = "base64"} | ConvertTo-Json) -ErrorAction Stop
            Write-Verbose "Uploaded image to Imgur: $($imgurResponse.Data.link)"
        }
        Catch {
            Throw $_.Exception
        }

        $newRowHash = @{PlexPath = $PlexPath; ImgurLink = $($imgurResponse.data.link); DeleteHash = $($imgurResponse.data.deletehash); DateAdded = $(Get-Date -Format 'yyyy-MM-dd'); Height = $imgurResponse.data.height; Width = $imgurResponse.data.width}
        $newRow = New-Object PsObject -Property $newRowHash
        Export-CSV $FilePath -InputObject $newRow -Append -Force -NoTypeInformation

        Return $newRow
    }
}

#endregion

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

Try {
    $imgurCreds = Get-StoredCredential -Target imgur -ErrorAction Stop
    if (-not $imgurCreds) {
        Throw "No imgur credentials found."
    }
    Write-Verbose "Retrieved imgur credentials."
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

$style = @"
<html>
<head>
<style type="text/css">
table
    {
        width: 95%
    }
td
    {
        padding: 5px;
        vertical-align:middle;
    }
a
    {
        color: #1a4b7f;
        text-decoration: none;
    }
a:hover
    {
        text-decoration: underline;
    }
a:visited
    {
        color: #636;
    }
.center
    {
        text-align: center;
    }
ul
    {
        list-style-position: outside;
    }
h1
    { 
        display: block;
        font-size: 2em;
        margin-top: 0.67em;
        margin-bottom: 0.67em;
        margin-left: 0;
        margin-right: 0;
        font-weight: bold;
    }
h2
    { 
        display: block;
        font-size: 1.5em;
        margin-top: 0.83em;
        margin-bottom: 0.83em;
        margin-left: 0;
        margin-right: 0;
        font-weight: bold;
    }
@media (max-width: 768px)
    {
      table {
        width: 100%;
      }
      td {
        display: block;
        width: 100%;
      }
      .center {
          display: block;
          margin-left: auto;
          margin-right: auto;
          width: 50%;
      }
      a {
        text-decoration: underline;
      }
    }
</style>
</head>
<body>
"@
#endregion

$response = Invoke-RestMethod "$PlexUrl`:$PlexPort/library/recentlyAdded/?X-Plex-Token=$plexToken" -Headers @{"accept"="application/json"} | Select-Object -ExpandProperty MediaContainer | Select-Object -ExpandProperty Metadata
$response = $response | Where-Object {$_.addedAt -gt $startDate -and $_.librarySectionTitle -notin $ExcludeLib}
# Grab those libraries

$movies = $response |
    Where-Object {$_.type -eq 'movie'} |
    Sort-Object addedAt

$tvShows = $response |
    Where-Object {$_.type -eq 'season'} |
    Group-Object parentTitle

# Initialize the counters and lists
$movieCount = 0
$movieList = "<h2>Movies:</h2>"

if ($($movies | Measure-Object).count -gt 0) {
    foreach ($movie in $movies) {
        # TMDB rate-limits. Currently 40 requests every 10 seconds, so sleep 10 seconds every 15 movies to be safe (2 calls per movie)
        # https://developers.themoviedb.org/3/getting-started/request-rate-limiting
        if ($movieCount -gt 1 -AND $movieCount%15 -eq 0) {
            Write-Verbose "On movie $movieCount, waiting 10 seconds..."
            Start-Sleep -Seconds 10
        }
        $posterFail = $false
        $movieCount++

        Write-Verbose "Looking up $($movie.title) ($movieCount / $($movies | Measure-Object | Select-Object -ExpandProperty Count))."

        # Retrieve movie info from The Movie Database
        $simpleResponse = (Invoke-RestMethod "$searchURL/$($imdbIDformat.Matches($movie.guid).value)?api_key=$tmdbToken&language=en-US&external_source=imdb_id").movie_results

        # Assuming we have a valid response, pull detailed info on the movie
        if ($simpleResponse.id) {
            $detailedResponse = (Invoke-RestMethod "https://api.themoviedb.org/3/movie/$($simpleResponse.id)?api_key=$tmdbToken&language=en-US")
        }

        if ($detailedResponse.Title) {
            if ($UploadPostersToImgur) {
                Try {
                    Invoke-WebRequest -Uri "$PlexUrl`:$PlexPort$($movie.thumb)?X-Plex-Token=$plexToken" -OutVariable moviePoster -ErrorAction Stop
                    $imgurLink = Invoke-ImgurUpload -ClientID $imgurCreds.UserName -ImageBase64 ([convert]::ToBase64String($moviePoster.content)) -PlexPath $movie.thumb -FilePath $ImgurFilePath -ErrorAction Stop
                    $movieList += "<tr><td><img src=`"$($imgurLink.ImgurLink)`" width=154px height=$(154/$($imgurLink.width)*$($imgurLink.height)) class=`"center`"></td>"
                }
                Catch {
                    $posterFail = $true
                }
            }
            else {
                if ($detailedResponse.poster_path) {
                    $movieList += "<tr><td><img src=`"https://image.tmdb.org/t/p/w154$($detailedResponse.poster_path)`" class=`"center`"></td>"
                } else {
                    $posterFail = $true
                }
            }
            # If the poster was unavailable, substitute a Plex logo
            if ($posterFail) {$movieList += "<tr><td><img src=`"$imgPlex`" height=154px width=154px class=`"center`"></td>"}
            $movieList += "<td><b><a href=`"http://www.imdb.com/title/$($detailedResponse.imdb_ID)/`">$($detailedResponse.title)</a></b> ($(($detailedResponse.release_date).split('-')[0]))"
            $movieList += "<ul><li><i>Genre$(if($detailedResponse.Genres.count -gt 1){"s"}):</i> $($detailedResponse.Genres.name -join ', ')</li>"
            $movieList += "<li><i>Rating:</i> $($movie.contentRating)</li>"
            $movieList += "<li><i>Runtime:</i> $($detailedResponse.runtime) minutes</li>"
            $movieList += "<li><i>Star$(if($detailedResponse.role.count -gt 1){"s"}):</i> $($movie.role.tag -join ', ')</li>"
            $movieList += "<li><i>Plot:</i> $($movie.summary)</li>"
            $movieList += "<li><i>IMDB rating:</i> $(($detailedResponse.vote_average).toString("#.#"))/10</li>"
            $movieList += "<li><i>Added:</i> $(Get-Date $epoch.AddSeconds($movie.addedAt) -Format 'MMMM d')</li></ul></td>"
        }
        else {
            # If the movie couldn't be retrieved, fail gracefully
            $movieList += "<tr><td><img src=`"$imgPlex`" height=150px width=150px class=`"center`"></td><td><li>$($movie.title)</a> ($($movie.year)) - no additional information</li></td>"
        }
        $movieList += "</tr>"

        Clear-Variable simpleResponse, detailedResponse, posterFail
    }
    $movieList += "</table><br/>"
}

$tvCount = 0
$tvList = "<h2>TV Seasons:</h2>"

if ($($tvShows | Measure-Object).Count -gt 0) {
    foreach ($show in $tvShows) {
        # Sleep every 15 shows for that TMDB rate limiting, even just the first one since we probably just did a bunch of movies.
        if ($tvCount%15 -eq 0) {
            Write-Verbose "On show $tvCount, waiting 10 seconds..."
            Start-Sleep -Seconds 10
        }

        # Count it!
        $tvCount++
        $posterFail = $false

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
            if ($UploadPostersToImgur) {
                Try {
                    Invoke-WebRequest -Uri "$PlexUrl`:$PlexPort$($show.group.parentThumb | Select-Object -First 1)?X-Plex-Token=$plexToken" -OutVariable showPoster -ErrorAction Stop
                    $imgurLink = Invoke-ImgurUpload -ClientID $imgurCreds.UserName -ImageBase64 ([convert]::ToBase64String($showPoster.content)) -PlexPath $($show.group.parentThumb | Select-Object -First 1) -FilePath $ImgurFilePath -ErrorAction Stop
                    $tvList += "<tr><td><img src=`"$($imgurLink.ImgurLink)`" width=154px height=$(154/$($imgurLink.width)*$($imgurLink.height)) class=`"center`"></td>"
                }
                Catch {
                    Write-Verbose "Failed to get Imgur link for this show's poster. Exception: $($_.Exception.Message)"
                    $posterFail = $true
                }
            }
            else {
                if ($detailedResponse.poster_path) {
                    $tvList += "<tr><td><img src=`"https://image.tmdb.org/t/p/w154$($detailedResponse.poster_path)`" class=`"center`"></td>"
                } else {
                    $posterFail = $true
                }
            }
            # If the poster was unavailable, substitute a Plex logo
            if ($posterFail) {$tvList += "<tr><td><img src=`"$imgPlex`" height=154px width=154px class=`"center`"></td>"}

            $tvList += "<td><b><a href=`"http://www.imdb.com/title/$imdbID/`">$($show.name)</a></b> ($($detailedResponse.first_air_date.split('-')[0])-$($detailedResponse.last_air_date.split('-')[0]))"
            $tvList += "<ul>"
            if ($detailedResponse.genres) {$tvList += "<li><i>Genre:</i> $($detailedResponse.genres.name -join ", ")</li>"}
            if ($contentRating) {$tvList += "<li><i>Rating:</i> $contentRating</li>"}
            $tvList += "<li><i>Plot:</i> $($detailedResponse.overview)</li>"
            $tvList += "<li><i>Now available:</i><br/></li><ul>"
            foreach ($season in ($show.Group | Sort-Object @{e={$_.index -as [int]}})){
                $tvList += "<li>$($season.title) - $($season.leafCount) episode$(if ($season.leafCount -gt 1){"s"})</li>"
            }
        }
        else {
            # If the series couldn't be found in the DB, fail gracefully
            $tvList += "<tr><td><img src=`"$imgPlex`" height=150px width=150px class=`"center`"></td><td><li>$($show.name)</a></li>"
            $tvList += "<td><li><a href=`"http://www.imdb.com/title/$($omdbResponse.imdbID)/`">$($show.name)</a></li>"
            foreach ($season in $show.Group){
                $tvList += "<ul><li>$($season.title) - $($season.leafCount) episode$(if ($season.leafCount -gt 1){"s"})</li>"
            }
        }
        $tvList += "</ul></td>"

        Clear-Variable simpleResponse, detailedResponse, contentRating, imdbID, season
    }
    $tvList += "</table><br/>"
}

$collectionInfo = "<h2>Featured Collection:</h2>"

if ($FeatureRandomCollection) {
    $libraries = (Invoke-RestMethod "$PlexUrl`:$PlexPort/library/sections?X-Plex-Token=$plexToken").MediaContainer.Directory
    $movieLibraries = $libraries | Where-Object {$_.type -eq 'movie' -and $_.title -notin $ExcludeLib}
    
    $collections = @()
    foreach ($library in $movieLibraries) {
        $collections += (Invoke-RestMethod -Uri "$PlexUrl`:$PlexPort/library/sections/$($library.key)/all?type=18&context=content.collections&X-Plex-Token=$plexToken").MediaContainer.Directory
    }

    if ($collections) {
        $collection = $collections[(Get-Random -Minimum 0 -Maximum ($collections | Measure-Object).count)]

        if ($UploadPostersToImgur) {
            Invoke-WebRequest -Uri "$PlexUrl`:$PlexPort$($collection.thumb)?X-Plex-Token=$plexToken" -OutVariable collectionPoster
            $imgurLink = Invoke-ImgurUpload -ClientID $imgurCreds.UserName -ImageBase64 ([convert]::ToBase64String($collectionPoster.content)) -PlexPath $collection.thumb -FilePath $ImgurFilePath
            $collectionInfo += "<tr><td><img src=`"$($imgurLink.ImgurLink)`" width=154px height=$(154/$($imgurLink.width)*$($imgurLink.height)) class=`"center`"></td>"
        }
        else {
            # If the poster wasn't uploaded, substitute a Plex logo
            $collectionInfo += "<tr><td><img src=`"$imgPlex`" height=154px width=154px class=`"center`"></td>"
        }
        $collectionInfo += "<td><b>$($collection.title)</b> ($($collection.minYear)$(if ($collection.MaxYear -gt $collection.MinYear){" - $($collection.MaxYear)"}))"
        $collectionInfo += "<ul>"
        if ($collection.summary) {
            $collectionInfo += "<li><i>Information:</i> $($collection.summary -replace '\n',' ')</li>"
        }
        $collectionInfo += "<li><i>Movie Count:</i> $($collection.childCount.ToString())</li>"
        $collectionInfo += "<li><i>Last Updated:</i> $(Get-Date $epoch.AddSeconds($collection.updatedAt) -Format 'MMMM d')</li></ul></td>"
        $collectionInfo += "</table><br/>"
    }
    
}


if (($movieCount -eq 0) -AND ($tvCount -eq 0)) {
    $body = "$style`Sorry, but no movies or TV shows have been added to the Plex library in the past $days days. Check out the Featured Collection in the meantime!<br/><br/>"

    if ($collection) {
        $body += $collectionInfo -replace '\n','<br>'
    }
} else {
    $body = "$style`<h1>Hello!</h1><br/>Here's the list of additions to my Plex library in the past $days days.<br/>"

    if ($movieCount -gt 0) {
        $body += $movieList -replace '\n','<br>'
    }

    if ($tvCount -gt 0) {
        $body += $tvList -replace '\n','<br>'
    }

    if ($collection) {
        $body += $collectionInfo -replace '\n','<br>'
    }

    $body += "Enjoy!"
}



if (-not $OmitVersionNumber) {
    $body += "<br><br><br><br><p align = right><font size = 1 color = Gray>Plex Version: $((Invoke-RestMethod "$PlexUrl`:$PlexPort/?X-Plex-Token=$plexToken" -Headers @{"accept"="application/json"}).mediaContainer.version). Posters/metadata from TMDb.</p></font>"
}

$body += "</body></html>"

$startDate = Get-Date (Get-Date).AddDays(-$days) -Format 'MMM d'
$endDate = Get-Date -Format 'MMM d'

    
# If not otherwise specified, set the To address the same as the From
if ($EmailTo -eq 'default') {
    $EmailTo = $emailCreds.UserName
}
$subject = "Plex Additions from $startDate-$endDate"

if (-not($PreventSendingEmptyList -and (($movieCount+$tvCount) -eq 0))) {
    Send-MailMessage -From $($emailCreds.UserName) -to $EmailTo -SmtpServer $SMTPserver -Port $SMTPport -UseSsl -Credential $emailCreds -Subject $subject -Body $body -BodyAsHtml -Encoding UTF8
    Write-Verbose "Sent email to $EmailTo."
}
