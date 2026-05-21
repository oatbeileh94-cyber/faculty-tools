@echo off
:: Run this once as Administrator to register the auto-start task
schtasks /create /tn "FlaskWebApp" /tr "\"%~dp0run.bat\"" /sc ONLOGON /ru "%USERNAME%" /rl HIGHEST /f
echo Task created. The Flask app will now start automatically on login.
echo If it crashes, restart it manually by running run.bat or rebooting.
pause
