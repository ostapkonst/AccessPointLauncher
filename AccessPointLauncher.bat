@setlocal EnableExtensions EnableDelayedExpansion
@echo off

:: title: скрипт для управления WiFi-точкой доступа (AP)
:: description: манипуляции с AP: просмотр статуса, запуск, перезапуск, остановка, автозапуск, расшаривание интернета
:: author: Константинов О. В.

:: параметры запуска по умолчанию
:: -mode=help -pause_timeout=20 -check_admin=yes -wait_services=yes
:: -check_hostednetwork=yes -paused_exit=no -logging=no

set mode=help
set pause_timeout=20
set check_admin=yes
set wait_services=yes
set check_hostednetwork=yes
set paused_exit=no
set logging=no
set latest_log=yes
set run_installed=no
set clear_log=no

:: не рекомендуется переопределять
set status=failed
set session=external
set english_codepage=437
set script_name=AccessPointLauncher
set task_title=%script_name%
set self_path="%~f0"
set log_file="%temp%\soft_ap-%date%.txt"
set system_log_file="%windir%\Temp\soft_ap-%date%.txt"
set task_file="%temp%\%script_name%.xml"
set vbs_file="%temp%\%script_name%.vbs"
set script_file="%windir%\%script_name%.bat"
set script=cmd /c %self_path% %* -session=internal
:: должен быть всегда равен 0
set /a bat_ok_exit=0
set /a bat_error_exit=1

:: можно переопределить параметры
:read_cmd_params
if not %1/==/ (
	if not "%__var%"=="" (
		if not "%__var:~0,1%"=="-" (
			endlocal
			goto read_cmd_params
		)
		endlocal & set %__var:~1%=%~1
	) else (
		setlocal & set __var=%~1
	)
	shift
	goto read_cmd_params
)

for /f "tokens=2 delims=:" %%i in ('chcp') do set prev_encoding=%%i
set prev_encoding=%prev_encoding: =%

if /i "%session%" == "external" (
	(goto) 2>nul & (
		chcp %english_codepage% >nul 2>&1

		if /i "%logging%" == "yes" (
			echo LogFile: %log_file%
			%script% >> %log_file% 2>&1
		) else (
			%script%
		)

		if errorlevel 1 (
			chcp %prev_encoding% >nul 2>&1
			if /i "%paused_exit%" == "yes" pause
			cmd /c exit /b %bat_error_exit%
		) else (
			chcp %prev_encoding% >nul 2>&1
			if /i "%paused_exit%" == "yes" pause
			cmd /c exit /b %bat_ok_exit%
		)
	)
)

if /i "%logging%" == "yes" (
	echo.=========begin=========
	echo.%date%:%time%
	echo.Mode:        %mode%
	echo.Timeout:     %pause_timeout%
	echo.Permissions: %check_admin%
	echo.Services:    %wait_services%
	echo.Support:     %check_hostednetwork%
	echo.Paused:      %paused_exit%
	echo.Logging:     %logging%
	echo.Latest:      %latest_log%
	echo.Execute:     %run_installed%
	echo.Clear:       %clear_log%
	echo.=======================
)

call :validate_input_params invalid_params invalid_list
if not "%invalid_params%" == "0" (
	echo Parameter "%invalid_list%" is not valid.
	echo Run .bat with -mode=help for getting help.

	exit /b %bat_error_exit%
)

if /i "%mode%" == "help" (
	echo Usage: AccessPointLauncher [-mode] [-pause_timeout] [-check_admin]
	echo             [-wait_services] [-check_hostednetwork] [-paused_exit]
	echo             [-logging] [-latest_log] [-run_installed] [-clear_log]
	echo.
	echo Default: AccessPointLauncher -mode=help -pause_timeout=20
	echo                              -check_admin=yes -wait_services=yes
	echo                              -check_hostednetwork=yes -paused_exit=no
	echo                              -logging=no -latest_log=yes -run_installed=no
	echo                              -clear_log=no
	echo.
	echo Options:
	echo     -mode=[help^|status^|start^|stop^|fresh^|restart^|install^|uninstall^|shared^|logs^]
	echo     -pause_timeout=^<positive number^>
	echo     -check_admin=[yes^|no]
	echo     -wait_services=[yes^|no]
	echo     -check_hostednetwork=[yes^|no]
	echo     -paused_exit=[yes^|no]
	echo     -logging=[yes^|no]
	echo     -latest_log=[yes^|no]
	echo     -run_installed=[yes^|no]
	echo     -clear_log=[yes^|no]

	exit /b %bat_ok_exit%
)

if /i "%mode%" == "logs" (
	if /i "%logging%" == "yes" (
		echo CurrentLogFile: %log_file%
	) else (
		call :print_log_file %system_log_file% %latest_log%
		if /i not %log_file% == %system_log_file% (
			call :print_log_file %log_file% %latest_log%
		)

		set nothing_to_show=yes
		if exist %system_log_file% set nothing_to_show=no
		if exist %log_file% set nothing_to_show=no
		if "!nothing_to_show!" == "yes" (
			echo Logs files is not found.
		) else (
			if /i "%clear_log%" == "yes" (
				del /f %system_log_file% >nul 2>&1
				del /f %log_file% >nul 2>&1
			)
		)
	)

	exit /b %bat_ok_exit%
)

where cmd.exe >nul 2>&1
if errorlevel 1 (
	echo This version of Windows is not support.
	goto exit
)

if not "%prev_encoding%" == "%english_codepage%" (
	echo Command line has incorrectly codepage, need %english_codepage%.
	goto exit
)

if /i "%mode%" == "install" goto check_permissions
if /i "%mode%" == "uninstall" goto check_permissions

:: ждем запуска сетевых служб
:starting_network_service
if /i "%wait_services%" == "yes" (
	echo | set /p="Waiting for the WLAN AutoConfig service: "

	:loop_service
	sc query wlansvc | find "RUNNING" > nul
	if errorlevel 1 (
		if /i "%logging%" == "yes" (
			echo Failure.
			goto exit
		) else (
			timeout /t 5 /nobreak > nul
			goto loop_service
		)
	) else (
		echo Success.
		echo.
	)
)

:: проверка поддержки точки доступа
:check_hostednetwork_support
if /i "%check_hostednetwork%" == "yes" (
	echo | set /p="WiFi is supported hosted network: "
	netsh wlan show drivers | findstr /e /r /c:"Hosted network supported.*Yes" > nul
	if errorlevel 1 (
		echo Failure.
		goto exit
	) else (
		echo Success.
		echo.
	)
)

if /i "%mode%" == "status" goto work_with_hostednetwork
if /i "%mode%" == "stop" goto work_with_hostednetwork

:: проверка на наличие прав администратора
:check_permissions
if /i "%check_admin%" == "yes" (
	echo | set /p="Have administrative permissions: "
	net session >nul 2>&1
	if errorlevel 1 (
		echo Failure.
		goto exit
	) else (
		echo Success.
		echo.
	)
)

:: проверяем тип запуска скрипта
:work_with_hostednetwork
if /i "%mode%" == "install" (
	echo Creating a task to run the script automatically...
	echo.
	goto install
) else if /i "%mode%" == "uninstall" (
	echo Deleting task and script from windows directory...
	echo.
	goto uninstall
) else if /i "%mode%" == "restart" (
	echo Forcibly launch the access point...
	echo.
	goto startup_hostednetwork
) else if /i "%mode%" == "shared" (
	echo Setting options...
	echo.
	goto shared_with_ics
)

netsh wlan show hostednetwork | findstr /e "Started" > nul
if errorlevel 1 (
	netsh wlan show hostednetwork | findstr /e "configured>" > nul
	if errorlevel 1 (
		echo The access point is not active.

		if /i "%mode%" == "status" (
			call :show_settings _
		) else (
			set start_AP=no
			if /i "%mode%" == "start" set start_AP=yes
			if /i "%mode%" == "fresh" set start_AP=yes
			if "!start_AP!" == "yes" (
				echo.
				echo Launch the access point...
				echo.
				goto startup_hostednetwork
			)
		)
	) else (
		echo The access point is not configured.
	)
) else (
	echo Access point already running.
	echo.

	netsh wlan show hostednetwork | findstr /e /r /c:"Number.*: 0" > nul
	if errorlevel 1 (
		echo There are active access point users.
		if /i "%mode%" == "status" (
			call :show_status
		)
	) else (
		echo No active access point users.
		if /i "%mode%" == "status" (
			call :show_settings _
		) else if /i "%mode%" == "fresh" (
			echo.
			echo Refresh the access point...
			echo.
			goto startup_hostednetwork
		)
	)

	if /i "%mode%" == "stop" (
		echo.
		netsh wlan stop hostednetwork > nul
		if errorlevel 1 (
			echo Perhaps access point is not stoped.
		) else (
			set status=success
			echo Access point stoped.
		)
		goto exit
	)
)

set status=success
goto exit

:validate_input_params
setlocal
set /a "invalid_params=0"
set "params_list="

set "is_valid=no"
for %%i in (help status start stop fresh restart install uninstall shared logs) do (
	if /i "%mode%" == "%%i" set "is_valid=yes"
)
if "%is_valid%" == "no" (
	set /a "invalid_params+=1"
	set "params_list[!invalid_params!]=mode"
)

set "is_valid=no"
for /l %%i in (1, 1, 300) do if /i "%pause_timeout%" == "%%i" set "is_valid=yes"
if "%is_valid%" == "no" (
	set /a "invalid_params+=1"
	set "params_list[!invalid_params!]=pause_timeout"
)

for %%k in (check_admin wait_services check_hostednetwork paused_exit logging latest_log run_installed clear_log) do (
	set "is_valid=no"
	for %%i in (yes no) do if /i "!%%k!" == "%%i" set "is_valid=yes"
	if "!is_valid!" == "no" (
		set /a "invalid_params+=1"
		set "params_list[!invalid_params!]=%%k"
	)
)

for /l %%i in (1, 1, %invalid_params%) do (
	if %%i == 1 (
		set "params_list=%params_list[1]%"
	) else (
		set "params_list=!params_list!, !params_list[%%i]!"
	)
)
endlocal & set "%1=%invalid_params%" & set "%2=%params_list%"

exit /b

:print_log_file
setlocal
set "cur_log_file=%~1"
set "cur_latest_log=%~2"
if exist %cur_log_file% (
	set /a "log_file_pos=0"
	if /i "%cur_latest_log%" == "yes" (
		for /f "delims=:" %%a in ('findstr /r /n /b "=*begin" %cur_log_file%') do (
			set /a "log_file_pos=%%a-1"
		)
	)
	echo LogFile: %cur_log_file%
	more +!log_file_pos! %cur_log_file%
)
endlocal

exit /b

:get_interface_from_ip
setlocal
set "interface_name="
set "interface_number="
set "local_ip=%~1"
set "_adapter="
set "_ip="
set /a "_number=0"
for /f "tokens=1* delims=:" %%g in ('ipconfig /all') do (
	set "_tmp=%%g"
	if "!_tmp:adapter=!"=="!_tmp!" (
		if not "!_tmp:IPv4 Address=!"=="!_tmp!" (
			set "_ip=%%h" & set "_ip=!_ip: =!"
			for /f "delims=(" %%k in ("!_ip!") do set "_ip=%%k"
			if "!_ip!"=="%local_ip%" (
				set "interface_name=!_adapter!"
				set "interface_number=!_number!"
			)
		)
	) else (
		set "_ip="
		set "_adapter=!_tmp:*adapter =!"
		set /a "_number+=1"
	)
)
endlocal & set "%2=%interface_name%" & set "%3=%interface_number%"

exit /b

:get_guid_from_interface
setlocal
set "guid="
set "adapter=%~1"
set "interface_number=%~2"
set "_adapterfound=false"
set /a "_number=0"
for /f "tokens=1* delims=:" %%f in ('netsh trace show interfaces') do (
	set "item=%%f"
	if not "!item:adapter=!"=="!item!" (
		set /a "_number+=1"
		if "%interface_number%"=="!_number!" (
			set "_adapterfound=true"
		)
	) else if not "!item!"=="!item:GUID=!" if "!_adapterfound!"=="true" (
		set "guid=%%g" & set "guid=!guid: =!"
		set "_adapterfound=false"
	)
)
endlocal & set "%3=%guid%"

exit /b

:: показываем имя и пароль точки доступа
:show_settings
setlocal
set name=
set channel=
set clients=
set mac_address=
set ip_address=
set password=

echo.
echo Access point settings:
for /f "tokens=1* delims=:" %%f in ('netsh wlan show hostednetwork') do (
	set item=%%f & set item=!item: =!
	if /i "!item!"=="SSIDname" (
		set name=%%g & set name=!name:~1!
		echo Name: !name!
	) else if /i "!item!"=="Channel" (
		set channel=%%g & set channel=!channel:~1!
		echo Channel: !channel!
	) else if /i "!item!"=="Numberofclients" (
		set clients=%%g & set clients=!clients:~1!
		echo Connected: !clients!
	) else if /i "!item!"=="BSSID" (
		set mac_address=%%g & set mac_address=!mac_address: =!
		set mac_address=!mac_address::=-!
		rem echo MAC-address: !mac_address!

		for /f "tokens=1* delims=:" %%i in ('ipconfig /all') do (
			if /i "%%j"==" !mac_address!" (
				set adapterfound=true
			) else if "!adapterfound!"=="true" (
				echo %%i | find "IPv4 Address" > nul
				if not errorlevel 1 (
					set ip_address=%%j & set ip_address=!ip_address: =!
					for /f "delims=(" %%k in ("!ip_address!") do set ip_address=%%k
					set adapterfound=false
					echo IP-address: !ip_address!

					echo | set /p="Internet: "
					ping -S "!ip_address!" www.google.com -n 1 -w 1000 >nul 2>&1
					if errorlevel 1 (
						echo No
					) else (
						echo Yes
					)
				)
			)
		)
	)
)

set pass_cmd="netsh wlan show hostednetwork security | findstr /r /c:"User.*key *:""
for /f "tokens=2 delims=:" %%k in ('%pass_cmd%') do (
	set password=%%k & set password=!password:~1!
	echo Password: !password!
)

set settings=no
if defined name set settings=yes
if defined channel set settings=yes
if defined clients set settings=yes
if defined ip_address set settings=yes
if defined password set settings=yes
if "%settings%" == "no" (
	echo.
	echo Access point not configured.
)

endlocal & set "%1=%ip_address%"

exit /b

:: показываем список ip-адресов подключенных клиентов
:show_status
setlocal
call :show_settings ip_address
echo.
set client_mac_address=
echo With the following connected clients:
if defined ip_address (
	set arp_params=-a -N "%ip_address%"
) else (
	set arp_params=-a
)

for /f "skip=16" %%k in ('netsh wlan show hostednetwork') do (
	set client_mac_address=%%k
	set client_mac_address=!client_mac_address::=-!

	set client_ip_address=
	set show_warning=no
	for /f "tokens=1" %%i in ('arp %arp_params% ^| find /i "!client_mac_address!"') do (
		if not defined client_ip_address (
			set client_ip_address=%%i
			echo | set /p=ip: !client_ip_address! 
			for /f "tokens=1-3" %%a in ('ping -n 1 -i 1 -a "!client_ip_address!"') do (
				if "%%c" == "[!client_ip_address!]" (
					set host_name=%%b
					echo | set /p=- !host_name! 
				)
			)
		) else (
			if "!show_warning!"=="no" (
				echo | set /p=*
				set show_warning=yes
			)
		)
	)

	if not defined client_ip_address (
		echo | set /p=mac: !client_mac_address!
	)
	echo.
)

if not defined client_mac_address (
	echo.
	echo The access point has no users.
)

endlocal

exit /b

:: настройка общего доступа подключения к интернету
:shared_with_ics
call :show_settings ip_address > nul

set private_int=
set private_guid=
if defined ip_address (
	echo Private ip-address: !ip_address!
	call :get_interface_from_ip !ip_address! private_int interface_number

	if defined private_int (
		echo Private interface: "!private_int!"
		call :get_guid_from_interface "!private_int!" !interface_number! private_guid
		if defined private_guid echo Private GUID: !private_guid!
	)
) else (
	echo Could not find private ip-address.
)

set metric=
set public_ip=
for /f "tokens=4*" %%f in ('route print ^| findstr "\<0.0.0.0\>"') do (
	if "%%g" neq "" (
		if not defined metric set /a metric=%%g
		if %%g leq !metric! (
			set public_ip=%%f
			set /a metric=%%g
		)
	)
)

echo.
set public_int=
set public_guid=
if defined public_ip (
	echo Public ip-address: %public_ip%
	set local_ip=%public_ip%
	call :get_interface_from_ip !public_ip! public_int interface_number

	if defined public_int (
		echo Public interface: "!public_int!"
		call :get_guid_from_interface "!public_int!" !interface_number! public_guid
		if defined public_guid echo Public GUID: !public_guid!
	)
) else (
	echo Could not find public ip-address.
)

echo.
if defined private_guid (
	if defined public_guid (
		if /i not "%private_guid%" == "%public_guid%" (
			del /f %vbs_file% >nul 2>&1

			set copy=no
			for /f "tokens=* usebackq" %%a in (%self_path%) do (
				if "%%a" == ":ics_connections_script" (
					set copy=yes
				) else if "!copy!" == "yes" (
					echo %%a >> %vbs_file%
				)
				if "%%a" == "Main( )" set copy=no
			)

			if exist %vbs_file% (
				cscript //nologo %vbs_file% "%public_guid%" "%private_guid%"
				echo.
				if errorlevel 1 (
					echo Internet distribution set up failed.
				) else (
					echo Internet distribution set up successfully.
					set status=success
				)
			) else (
				echo Common connection configuration script file not found.
			)
		) else (
			echo Interface GUIDs must not match.
		)
	) else (
		echo Failed to get public GUID.
	)
) else (
	echo Failed to get private GUID.
)

goto exit 

:: запускаем точку доступа
:startup_hostednetwork
if not defined iteration (
	set pause=timeout /t "%pause_timeout%" /nobreak ^> nul

	set iteration=first
	if /i "%mode%" == "restart" set iteration=restart
)

if "%iteration%" == "first" (
	netsh wlan show hostednetwork | findstr /e "Started" > nul
	if errorlevel 1 (
		netsh wlan show hostednetwork | findstr /e "Allowed" > nul
		if errorlevel 1 (
			netsh wlan set hostednetwork allow
			if not errorlevel 1 %pause%
		)
	) else (
		netsh wlan stop hostednetwork
		if not errorlevel 1 %pause%
	)
) else (
	if "%iteration%" == "restart" (
		set iteration=first
	)
	netsh wlan set hostednetwork disallow
	if not errorlevel 1 (
		%pause%
		netsh wlan set hostednetwork allow
		if not errorlevel 1 %pause%
	)
)
if not errorlevel 1 (
	netsh wlan start hostednetwork
	if not errorlevel 1 (
		echo Access point started.
		set status=success
	)
)

if "%status%" == "failed" (
	if not "%iteration%" == "third" (
		echo Unable to start the AP on the %iteration% attempt. Try in a minute.
		echo.
		timeout /t 60 > nul
		if "%iteration%" == "first" (
			set iteration=second
		) else (
			set iteration=third
		)
		goto startup_hostednetwork
	)
	echo Perhaps there is no such access point.
)

if /i "%logging%" == "yes" (
	%pause%
	netsh wlan show hostednetwork

	echo.----------------------
	echo.%date%:%time%
	echo.Status: %status%
	echo.Attempt: %iteration%
	echo.----------------------
)

:: выходим обратно в консоль
:exit
echo.

if "%status%" == "success" (
	echo Script completed successfully.
	exit /b %bat_ok_exit%
) else (
	echo Script failed.
	exit /b %bat_error_exit%
)

:install
del /f %task_file% >nul 2>&1

set copy=no
for /f "tokens=* usebackq" %%a in (%self_path%) do (
	if "%%a" == ":install_task" (
		set copy=yes
	) else if "!copy!" == "yes" (
		echo %%a >> %task_file%
	)
	if "%%a" == "</Task>" set copy=no
)

if exist %task_file% (
	schtasks /create /f /xml %task_file% /tn %task_title% >nul 2>&1
	if errorlevel 1 (
		echo Failed to create access point startup task.
	) else (
		if /i %self_path% == %script_file% (
			echo The task to automatically start the access point was created successfully.
			set status=success
		) else (
			copy /y %self_path% %script_file% >nul 2>&1
			if errorlevel 1 (
				echo Could not create script file in windows directory.
			) else (
				echo The task to automatically start the access point was created successfully.
				set status=success
			)
		)
	)
) else (
	echo Task installation script file not found.
)

if "%status%" == "success" (
	if /i "%run_installed%" == "yes" (
		echo.
		echo | set /p="Try to execute task: "
		schtasks /run /tn %task_title% >nul 2>&1
		if errorlevel 1 (
			echo Failure.
			set status=failed
		) else (
			echo Success.
		)
	)
)

goto exit

:uninstall
set check_exist=schtasks /query /tn %task_title% 
%check_exist% >nul 2>&1
if errorlevel 1 (
	%check_exist% 2>&1 | findstr /e "Access is denied." > nul
	if errorlevel 1 (
		echo Could not delete task because it does not exist.
	) else (
		echo Could not delete task because not enough rights.
		goto exit
	)
) else (
	schtasks /delete /f /tn %task_title% >nul 2>&1
	if errorlevel 1 (
		echo Could not delete task.
		goto exit
	) else (
		echo Task successfully deleted.
	)
)

echo.
if exist %script_file% (
	if /i %script_file% == %self_path% (
		(goto) 2>nul & del /f %script_file% >nul 2>&1 & (
			if exist %script_file% (
				echo Could not delete script file.
				echo.
				echo Script failed.
				cmd /c exit /b %bat_error_exit%
			) else (
				echo Script file successfully deleted.
				echo.
				echo Script completed successfully.
				cmd /c exit /b %bat_ok_exit%
			)
		)
	) else (
		del /f %script_file% >nul 2>&1
		if exist %script_file% (
			echo Could not delete script file.
		) else (
			echo Script file successfully deleted.
			set status=success
		)
	)
) else (
	echo Script file does not exist in windows directory.
	set status=success
)

goto exit

:install_task
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.3" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Author>ostapkonst</Author>
    <Description>Runs the script when you turn on the computer or exit from sleep mode, which starts the access point. For more look: https://gist.github.com/ostapkonst/73343da2e9628c72ef898e9f235c5f4d</Description>
  </RegistrationInfo>
  <Triggers>
    <BootTrigger>
      <Enabled>true</Enabled>
      <Delay>PT1M</Delay>
    </BootTrigger>
    <EventTrigger>
      <Enabled>false</Enabled>
      <Subscription>&lt;QueryList&gt;&lt;Query Id="0" Path="Microsoft-Windows-Diagnostics-Performance/Operational"&gt;&lt;Select Path="Microsoft-Windows-Diagnostics-Performance/Operational"&gt;*[System[Provider[@Name='Microsoft-Windows-Diagnostics-Performance'] and EventID=300]]&lt;/Select&gt;&lt;/Query&gt;&lt;/QueryList&gt;</Subscription>
      <Delay>PT1M</Delay>
    </EventTrigger>
    <EventTrigger>
      <Enabled>false</Enabled>
      <Subscription>&lt;QueryList&gt;&lt;Query Id="0" Path="System"&gt;&lt;Select Path="System"&gt;*[System[Provider[@Name='Microsoft-Windows-Power-Troubleshooter'] and EventID=1]]&lt;/Select&gt;&lt;/Query&gt;&lt;/QueryList&gt;</Subscription>
      <Delay>PT1M</Delay>
    </EventTrigger>
    <EventTrigger>
      <Enabled>false</Enabled>
      <Subscription>&lt;QueryList&gt;&lt;Query Id="0" Path="Microsoft-Windows-Diagnostics-Performance/Operational"&gt;&lt;Select Path="Microsoft-Windows-Diagnostics-Performance/Operational"&gt;*[System[Provider[@Name='Microsoft-Windows-Diagnostics-Performance'] and EventID=100]]&lt;/Select&gt;&lt;/Query&gt;&lt;/QueryList&gt;</Subscription>
      <Delay>PT1M</Delay>
    </EventTrigger>
    <EventTrigger>
      <Enabled>false</Enabled>
      <Subscription>&lt;QueryList&gt;&lt;Query Id="0" Path="Microsoft-Windows-WLAN-AutoConfig/Operational"&gt;&lt;Select Path="Microsoft-Windows-WLAN-AutoConfig/Operational"&gt;*[System[Provider[@Name='Microsoft-Windows-WLAN-AutoConfig'] and EventID=8009]]&lt;/Select&gt;&lt;/Query&gt;&lt;/QueryList&gt;</Subscription>
      <Delay>PT1M</Delay>
    </EventTrigger>
    <EventTrigger>
      <Enabled>false</Enabled>
      <Subscription>&lt;QueryList&gt;&lt;Query Id="0" Path="System"&gt;&lt;Select Path="System"&gt;*[System[Provider[@Name='Microsoft-Windows-Kernel-General'] and EventID=12]]&lt;/Select&gt;&lt;/Query&gt;&lt;/QueryList&gt;</Subscription>
      <Delay>PT1M</Delay>
    </EventTrigger>
    <EventTrigger>
      <Enabled>true</Enabled>
      <Subscription>&lt;QueryList&gt;&lt;Query Id="0" Path="System"&gt;&lt;Select Path="System"&gt;*[System[Provider[@Name='Microsoft-Windows-Kernel-General'] and EventID=1]]&lt;/Select&gt;&lt;/Query&gt;&lt;/QueryList&gt;</Subscription>
      <Delay>PT1M</Delay>
    </EventTrigger>
    <EventTrigger>
      <Enabled>false</Enabled>
      <Subscription>&lt;QueryList&gt;&lt;Query Id="0" Path="System"&gt;&lt;Select Path="System"&gt;*[System[Provider[@Name='Microsoft-Windows-WLAN-AutoConfig'] and EventID=4000]]&lt;/Select&gt;&lt;/Query&gt;&lt;/QueryList&gt;</Subscription>
      <Delay>PT1M</Delay>
    </EventTrigger>
    <BootTrigger>
      <Repetition>
        <Interval>PT30M</Interval>
        <StopAtDurationEnd>false</StopAtDurationEnd>
      </Repetition>
      <Enabled>false</Enabled>
      <Delay>PT30M</Delay>
    </BootTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>S-1-5-18</UserId>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>true</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>true</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <DisallowStartOnRemoteAppSession>false</DisallowStartOnRemoteAppSession>
    <UseUnifiedSchedulingEngine>false</UseUnifiedSchedulingEngine>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT13M</ExecutionTimeLimit>
    <Priority>7</Priority>
    <RestartOnFailure>
      <Interval>PT5M</Interval>
      <Count>2</Count>
    </RestartOnFailure>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>%windir%\AccessPointLauncher.bat</Command>
      <Arguments>-mode=start -pause_timeout=25 -check_admin=no -wait_services=no -check_hostednetwork=yes -paused_exit=no -logging=yes</Arguments>
    </Exec>
  </Actions>
</Task>

:ics_connections_script
Option Explicit

const ICSSHARINGTYPE_PUBLIC = 0
const ICSSHARINGTYPE_PRIVATE = 1

sub Main( )
	dim objArgs
	set objArgs = WScript.Arguments

	if objArgs.Count = 2 then
		EnableDisableICS objArgs(0), objArgs(1)
	else
		dim szMsg
		szMsg = "Please provide the GUID of the connection as the argument." & _
				vbNewLine & vbNewLine & "Usage:" & vbNewLine & _
				"       " & WScript.scriptname & " ""Public GUID"" ""Private GUID"""
		WScript.Echo(szMsg)
	end if
end sub

function IsAdmin()
	On Error Resume Next
	CreateObject("WScript.Shell").RegRead("HKEY_USERS\S-1-5-19\Environment\TEMP")
	if Err.Number = 0 then 
		IsAdmin = True
	else
		IsAdmin = False
	end if
	Err.Clear
	On Error goto 0
end function

sub EnableDisableICS(sPublicConnectionGuid, sPrivateConnectionGuid)
	On Error Resume Next
	dim bFound1, bFound2
	dim objShare, objEveryColl, objNetConn

	bFound1 = False
	bFound2 = False

	if uCase(sPrivateConnectionGuid) = uCase(sPublicConnectionGuid) then
		Wscript.Echo("ICS Private Guid is equal to ICS Public Guid")
		WScript.Quit(2)
	end if

	if not IsAdmin() then
		Wscript.Echo("You don't have admin rights")
		WScript.Quit(3)
	end if

	set objShare = Wscript.CreateObject("HNetCfg.HNetShare.1")
	if not IsObject(objShare) then
		Wscript.Echo("Unable to get the HNetCfg.HnetShare.1 object")
		WScript.Quit(4)
	else
		if IsNull(objShare.SharingInstalled) then
			Wscript.Echo("Sharing isn't available on this platform")
			WScript.Quit(5)
		else
			set objEveryColl = objShare.EnumEveryConnection
			if not IsObject(objEveryColl) then
				Wscript.Echo("Unable to get the connections enumerator")
				WScript.Quit(6)
			end if
		end if
	end if

	if objEveryColl.Count < 2 then
		Wscript.Echo("Less than 2 connections at node")
		WScript.Quit(7)
	end if

	for each objNetConn In objEveryColl
		dim objShareCfg
		set objShareCfg = objShare.INetSharingConfigurationForINetConnection(objNetConn)
		if IsObject(objShareCfg) then
			dim objNCProps
			set objNCProps = objShare.NetConnectionProps(objNetConn)

			if IsObject(objNCProps) then
				dim szMsg
				szMsg = "    Guid: "       & objNCProps.Guid & vbNewLine & _
						"    DeviceName: " & objNCProps.DeviceName & vbNewLine & _
						"    Status: "
				Err.Clear
				select case objNCProps.Guid
					case sPrivateConnectionGuid
						objShareCfg.EnableSharing(ICSSHARINGTYPE_PRIVATE)
						Wscript.Echo("Start ICS Private on connection:")
						if Err.Number = 0 then
							Wscript.Echo(szMsg & "Success")
							bFound1 = objShareCfg.SharingEnabled
						else
							Wscript.Echo(szMsg & "Failed")
						end if
						Wscript.Echo()
					case sPublicConnectionGuid
						objShareCfg.EnableSharing(ICSSHARINGTYPE_PUBLIC)
						Wscript.Echo("Start ICS Public on connection:")
						if Err.Number = 0 then
							Wscript.Echo(szMsg & "Success")
							bFound2 = objShareCfg.SharingEnabled
						else 
							Wscript.Echo(szMsg & "Failed")
						end if
						Wscript.Echo()
				end select
			end if
		end if
	next

	if bFound1 and bFound2 then
		Wscript.Echo("ICS sharing is successfully enabled")
		WScript.Quit(0)
	else
		Wscript.Echo("Unable to start the priv. or publ. connection")
		WScript.Quit(1)
	end if
end sub

Main( )
