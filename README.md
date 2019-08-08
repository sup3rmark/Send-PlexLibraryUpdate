# Send-PlexLibraryUpdate.ps1
A Powershell script to report on movies recently added to your Plex library. It uses the Plex API as well as OMDBapi.com (the Open Movie Database) to fetch information on each movie.

Requirements:
- Plex
- TVDB account and API token (https://www.themoviedb.org/settings/api)
- An email account and its SMTP server address/port (this is already filled in for Gmail, but to send from a Gmail account, "Access for less secure apps" must be turned *on*; instructions [here](https://support.google.com/accounts/answer/6010255?hl=en))
- CredentialManager module
- Tokens and SMTP email credentials stored in Windows Credential Manager

Instructions:

1. Download Send-PlexLibraryUpdate.ps1 to your computer from the link above.
2. Install the CredentialManager module (`Install-Module CredentialManager` should do the trick)
3. Open a Powershell prompt and run the command

        Set-ExecutionPolicy Unrestricted

4. Store the credentials of the email address you want to send the email from in Windows Credential Manager using the CredentialManager module:

```powershell
Install-Module CredentialManager
New-StoredCredential -Target plexToken -UserName plex -Password [Plex token] -Type Generic -Persist LocalMachine
New-StoredCredential -Target tmdb.org -UserName tmdb -Password [TMDB token] -Type Generic -Persist LocalMachine
New-StoredCredential -Target PlexCheck -UserName [Email address] -Password [Email password] -Type Generic -Persist LocalMachine
```
5. Run the script! You can run from a Powershell prompt, or by right-clicking and selecting *Run*.

Default behavior:

    .\Send-PlexLibraryUpdate.ps1

This will run the script with default values:

- It will assume you're running it from the same box that your Media Server is running on. As such, the default IP is 127.0.0.1.
- It uses the default Plex port, 32400.
- The sender email address will be what's defined in your credentials in Windows Credential Manager.
- The to address, unless otherwise specified, will be the same as the sender address. Change it at run-time.
- This defaults to Gmail's SMTP server settings, but that can be overridden with parameters.
- This will send over SSL, and this can't be overridden. Just do it.
- This defaults to a 7-day lookback period, but that can be changed, if that's what you're into.

Here's an example of how to run it with all the values changed:

    .\Send-PlexLibraryUpdate.ps1 -cred foo -url 10.0.0.100 -port 12345 -days 14 -emailTo 'test@email.com' -smtpServer 'smtp.server.com' -smtpPort 132

Please let me know if you have any questions, comments, or suggestions!
