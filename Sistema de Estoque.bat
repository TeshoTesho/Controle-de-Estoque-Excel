@Echo off
@Mode 48,18
@Title %~n0
Batbox /h 0


:data
rem Caminho
set /p caminho=<url.txt
rem Data Atual
set dia=%date:~0,2%
set mesn=%date:~3,2%
if %mesn%==01 set mes=Janeiro
if %mesn%==02 set mes=Fevereiro
if %mesn%==03 set mes=Marco
if %mesn%==04 set mes=Abril
if %mesn%==05 set mes=Maio
if %mesn%==06 set mes=Junho
if %mesn%==07 set mes=Julho
if %mesn%==08 set mes=Agosto
if %mesn%==09 set mes=Setembro
if %mesn%==10 set mes=Outubro
if %mesn%==11 set mes=Novembro
if %mesn%==12 set mes=Dezembro
set ano=%date:~6,4%
set FORMATO=%mesn% %mes% - %ano% - Historico de Balanco
rem Data Antiga
set /a mesn_old=%date:~3,2%-1
if %mesn_old%==1 set mes_old=Janeiro
if %mesn_old%==2 set mes_old=Fevereiro
if %mesn_old%==3 set mes_old=Marco
if %mesn_old%==4 set me_olds=Abril
if %mesn_old%==5 set mes_old=Maio
if %mesn_old%==6 set mes_old=Junho
if %mesn_old%==7 set mes_old=Julho
if %mesn_old%==8 set mes_old=Agosto
if %mesn_old%==9 set mes_old=Setembro
if %mesn_old%==10 set mes_old=Outubro
if %mesn_old%==11 set mes_old=Novembro
if %mesn_old%==0 set mes_old=Dezembro
 set /a ano_old=%date:~6,4%	
if %mesn_old%==0 set /a ano_old=%date:~6,4%-1 


set FORMATO_old=0%mesn_old% %mes_old% - %ano_old% - Historico de Balanco
rem Título
title %mes% de %ano% - Sistema de Estoque NLA
goto menu


:menu
cls
Call Button  15 3 " Abrir Planilha"  18 11 "  Sair  " 17 8 "  Config  " # Press
Getinput /m %Press% /h 70
	:: Check for the pressed button 
	if %errorlevel%==1 (goto open)
	if %errorlevel%==2 (exit)
	if %errorlevel%==3 (goto conn)
	goto Loop




	:erro
	rem Seção de ERRO
	cls
	color 0C
	echo ERRO! Funcao criar atalho Indisponível
	timeout 10	



	:atalho
	rem Seção Atalhoo
	cls
	IF EXIST "%caminho%\%FORMATO.lnk%" (
		rem echo Atalho já existe
		) ELSE (		
		rem echo Criando Atalho
		echo set WshShell = WScript.CreateObject("WScript.Shell") >> CreateShortcut.vbs
		echo strDesktop = WshShell.SpecialFolders("Desktop") >> CreateShortcut.vbs
		echo set oUrlLink = WshShell.CreateShortcut("%caminho%\%FORMATO%.lnk") >> CreateShortcut.vbs
		echo oUrlLink.TargetPath = "%caminho%\Historico\%FORMATO%.xlsm" >> CreateShortcut.vbs
		echo oUrlLink.Save >> CreateShortcut.vbs
		cscript CreateShortcut.vbs
		del CreateShortcut.vbs
		)



		:historico
		cls
		IF EXIST "%caminho%\Historico" (
		rem echo Historico já existe
		) ELSE (
		rem echo Criando Histórico
		md Historico
		)



		:planilha
		cls
		IF EXIST "Historico/%FORMATO%.xlsm" (
		rem echo Planilha já Existe
		) ELSE (	
		rem echo Criando Planilha
		copy "00 Branco.xlsm" "Historico\%FORMATO%.xlsm"
		)	


		:exe
		cls
		IF EXIST "%FORMATO%".exe (
		rem echo Exe já existe
		) ELSE (
		rem echo Criando Exe
		type part1.txt >>"%FORMATO%".cs
		echo "%caminho%\historico\%FORMATO%.xlsm";>>"%FORMATO%".cs
		type part3.txt>>"%FORMATO%".cs
		Title Compilando Arquivo CS
		pushd "%caminho%\v2.0.50727"
		csc /target:winexe /out:"%caminho%\%FORMATO%.exe" "%caminho%\%FORMATO%.cs"
		pushd %caminho%
		del /f /q "*.cs"
		)

	:conf
	cls
	rem Configurações
	attrib +s +h "%caminho%\*.exe"
	attrib +s +h "%caminho%\*.txt"
	attrib -s -h "%caminho%\*.bat"
	attrib +s +h "%caminho%\Button.bat"
	attrib -s -h "%caminho%\*.xlsm"


	:abrir
	cls
	rem Abrir Planilha
	call "%caminho%\%FORMATO%.exe"

	:sair
	cls
	rem Sair da Aplicação
	Exit

	:aberto
	cls
	echo Planilha já está aberta


	:excluir
	cls
	del /f /a "%caminho%\%FORMATO%.exe"
	del /f /a "%caminho%\historico\%FORMATO%.xlsm"
	cls
	goto menu

	:open
	cls
	rem Menu
	Title Abrindo Planilha...
	goto historico
	goto planilha
	goto exe
	goto abrir
	goto conf
	exit

	:open_old
	cls
	rem Abrir Planilha
	echo %FORMATO_old%>>teste.txt
	call "%caminho%\%FORMATO_old%.exe"
	exit


	:conn
	cls
	Call Button  12 3 " Abrir Planilha Antiga"  12 8 "Deletar Planilha Atual" 12 11 "Voltar" 30 11 "Sair" # Press
	Getinput /m %Press% /h 70
	:: Check for the pressed button 
	if %errorlevel%==1 (goto open_old)
	if %errorlevel%==2 (goto excluir)
	if %errorlevel%==3 (goto menu)
	if %errorlevel%==4 (exit)
	goto conn