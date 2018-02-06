'=========================================================================
'==== Requiere Psexec y nircmd
'==== mueve nircmd a equipo remoto y ejecuta con psexec en modo silencioso
'=========================================================================
@echo off
CLS

set IP=
echo  Introduce la IP del equipo:
set /p ip=%ip%

:CompruebaPing
ping %ip% -n 1 |find "inaccesible" 1>nul
if errorlevel 1 goto credenciales
if errorlevel 0 goto apagado

:credenciales
set usuario=
echo  usuario:
set /p usuario=%usuario%


set password=
echo password:
set /p password=%password%

:mapeo

net use s: \\%ip%\c$ /user:%usuario% %password%
if errorlevel 2 goto repara
if errorlevel 1 goto fin


xcopy nircmd.exe s:\Temp\ /I /S /Y

psexec \\%ip% -u %usuario% -p %password% -i C:\Temp\nircmd.exe 0 savescreenshot C:\Temp\screenshot.jpg

xcopy S:\Temp\screenshot.jpg screenshot.jpg /Y /I
del s:\Temp\screenshot.jpg
del s:\Temp\nircmd.exe

start screenshot.jpg

:fin

net use s: /d

exit

:error
echo  No se ha podido ejecutar el proceso automatico correctamente. Debe realizarse manualmente:
echo      1. Conectate por control remoto al equipo.
echo      2. Habilita manualmente los recursos ADMIN$ y C$.
echo      3. Vuelve a ejecutar "ejecuta.bat".
pause
exit

:: "Repara" el acceso remoto a equipos, generando los recursos compartidos ADMIN$ y C$ mediante WMIC
:repara
echo  Tratando de generar los recursos compartidos ADMIN$ y C$...
wmic /node:%ip% /user:"%usuario%" /password:"%password%" process call create "net share admin$"
wmic /node:%ip% /user:"%usuario%" /password:"%password%" process call create "net share c$=c:\"
if errorlevel 1 goto error
goto mapeo

:apagado
Echo El equipo %ip% esta apagado o no responde al ping.
pause
