# PlexCheck
A Powershell script to report on movies recently added to your Plex library. It uses the Plex API as well as OMDBapi.com (the Open Movie Database) to fetch information on each movie.

Requirements:
- Windows
- Powershell
- Credentials stored in Windows Credential Manager
- Get-CredentialFromWindowsCredentialManager.ps1 from Tobias Burger: https://gist.github.com/toburger/2947424

This should be relatively easy to set up for someone running their server on a Windows machine, but give me a second and I'll have a write-up for those who might not know what to do.

1. Download PlexCheck.ps1 to your computer from my link above.
2. Download [Get-CredentialFromWindowsCredentialManager.ps1 from Tobias Burger](https://gist.github.com/toburger/2947424)
3. Change the .ps1 extension on Get-CredentialFromWindowsCredentialManager.ps1 to .psm1, and save it to C:\Users\[username]\Documents\WindowsPowershell\Modules\Get-CredentialFromWindowsCredentialManager\.
4. Open a Powershell prompt and run the command

        Set-ExecutionPolicy Unrestricted

5. Store the credentials of the email address you want to send the email from in Windows Credential Manager. Instructions [here](http://windows.microsoft.com/en-us/windows7/store-passwords-certificates-and-other-credentials-for-automatic-logon). You should store them with the name "PlexCheck" to avoid having to specify a credential name when running the script.
6. Run the script! You can run from a Powershell prompt, or by right-clicking and selecting *Run*.

Default behavior:

    .\PlexCheck.ps1

This will run the script with default values:

- It will assume you're running it from the same box that your Media Server is running on. As such, the default IP is 127.0.0.1.
- It uses the default Plex port, 32400.
- The sender email address will be what's defined in your credentials in Windows Credential Manager.
- The to address, unless otherwise specified, will be the same as the sender address. Change it at run-time.
- This defaults to Gmail's SMTP server settings, but that can be overridden with parameters. To send from a Gmail account, "Access for less secure apps" must be turned *on*. Instructions [here](https://support.google.com/accounts/answer/6010255?hl=en).
- This will send over SSL, and this can't be overridden. Just do it.
- This defaults to a 7-day lookback period, but that can be changed, if that's what you're into.

Here's an example of how to run it with all the values changed:

    .\PlexCheck.ps1 -cred foo -url 10.0.0.100 -port 12345 -days 14 -emailTo 'test@email.com' -smtpServer 'smtp.server.com' -smtpPort 132

This is the first time I've publicly shared a script I wrote, so please let me know if you have any questions, comments, or suggestions!
