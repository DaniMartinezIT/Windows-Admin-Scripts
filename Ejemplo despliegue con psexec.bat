@echo off

:: -------------------------------------------------------------------------------------------------------------------------------------------------------
:: Añadir los valores correspondientes que se deseen en las variables script
:: este nombre se usara para generar un log al final del procedimiento, rutaInstalacion (ruta donde se copiara en el pc remoto),
:: ejecutable (fichero a sjecutar, puede ser un msi, un bat,...) y el comando que se lanzará
:: El ejecutable debe de encontrarse en la misma ruta que el script y el comando
:: -------------------------------------------------------------------------------------------------------------------------------------------------------

set IP=
echo  Introduce la IP del equipo:
set /p ip=%ip%

:CompruebaPing
ping %ip% -n 1 |find "inaccesible" 1>nul
if errorlevel 1 goto credenciales
if errorlevel 0 goto apagado

:credenciales
set usuario=0
echo  Introduce el nombre de usuario "Administrador" local (intro para usaurio por defecto):
set user=0
set /p user=%user%
if  %user%==0 (set usuario=administrador) ELSE (set usuario=%user%)

set password=0
echo  Introduce password del usuario "Administrador" local (intro para password por defecto):
set pass=0
set /p pass=%pass%
if  %pass%==0 (set password=predeterminada) ELSE (set password=%pass%)

:: -------------------------------------------------------------------------------------------------------------------------------------------------------
set script=flash_16_0_0_35
set rutaInstalacion=flash16
set ejecutable=install_flash_player_16_active_x.msi
set comando=psexec \\%ip% -u %usuario% -p %password% C:\Windows\System32\msiexec.exe /i C:\%rutaInstalacion%\%ejecutable% /qn /l*v C:\%rutaInstalacion%\%script%.log
:: -------------------------------------------------------------------------------------------------------------------------------------------------------


:mapeo
net use s: /d
echo  Mapeando unidad de usuario en "S:"...
net use s: \\%ip%\c$ /user:%usuario% %password%
if errorlevel 2 goto repara
if errorlevel 1 goto fin

echo  Copiando ficheros necesarios para la ejecucion...
xcopy %ejecutable% s:\%rutaInstalacion%\ /I /S /Y

echo  Instalando en el equipo.
%comando%

:fin
echo  Desmapeando unidad...
net use s: /d
echo.
echo Proceso finalizado correctamente. Contacta con el usuario
echo para comprobar el correcto funcionamiento.
echo.
pause
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
