ECHO OFF
CLS
:MENU
ECHO.
ECHO iDrac Scripts
ECHO.
ECHO 1 - Start iDrac Prerequisite
ECHO 2 - Start iDrac Configuration
ECHO 3 - Start iDrac PowerUp
ECHO 4 - Start iDrac Shutdown
ECHO 5 - Start iDrac Rac Reset
ECHO 6 - Start iDrac Clear Config
ECHO 7 - Start iDrac Validation
ECHO 8 - EXIT
ECHO.
SET /P M=Type your choice and press ENTER:
IF %M%==1 GOTO PRE
IF %M%==2 GOTO CONF
IF %M%==3 GOTO POWERUP
IF %M%==4 GOTO SHUTDOWN
IF %M%==5 GOTO RACRESET
IF %M%==6 GOTO CLEAR
IF %M%==7 GOTO VALIDATION
IF %M%==8 GOTO EXIT
:PRE
powershell; python '"C:\Users\Admin\Desktop\iDrac Configuration Automation\pre_iDrac.py"'
GOTO MENU

:CONF
powershell; python '"C:\Users\Admin\Desktop\iDrac Configuration Automation\iDrac.py"'
GOTO MENU

:POWERUP
powershell; python '"C:\Users\Admin\Desktop\iDrac Configuration Automation\iDrac-Powerup.py"'
GOTO MENU

:SHUTDOWN
powershell; python '"C:\Users\Admin\Desktop\iDrac Configuration Automation\iDrac-Shutdown.py"'
GOTO MENU

:RACRESET
powershell; python '"C:\Users\Admin\Desktop\iDrac Configuration Automation\iDrac-Racreset.py"'
GOTO MENU

:CLEAR
powershell; python '"C:\Users\Admin\Desktop\iDrac Configuration Automation\iDrac-Clear.py"'
GOTO MENU

:VALIDATION
powershell; python '"C:\Users\Admin\Desktop\iDrac Configuration Automation\iDrac-Validation.py"'
GOTO MENU

:EXIT
exit