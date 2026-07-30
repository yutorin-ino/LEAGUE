@echo off
rem Open grommet_1100_dashboard.html in Google Chrome
set "DASHBOARD=%~dp0grommet_1100_dashboard.html"

set "CHROME=%ProgramFiles%\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME%" set "CHROME=%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME%" set "CHROME=%LocalAppData%\Google\Chrome\Application\chrome.exe"

rem Use the Chrome profile signed in as yutorin.ino@gmail.com so Gmail links
rem always compose from that account, regardless of which machine/session opens this.
set "CHROME_PROFILE=Profile 1"

if exist "%CHROME%" (
    start "" "%CHROME%" --profile-directory="%CHROME_PROFILE%" "%DASHBOARD%"
) else (
    echo Google Chrome not found. Opening with default browser instead.
    start "" "%DASHBOARD%"
)
