@echo off
:data
set caminho=\\COL_PGRANDE_001\DOCUMENTOS\ESTOQUE

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
title %FORMATO%.xlsm


goto menu


:erro
	rem Seção de ERRO
	cls
	color 0C
	echo ERRO! Funcao criar atalho Indisponível
	timeout 10	



:atalho
	rem Seção Atalhoo
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
	IF EXIST "%caminho%\Historico" (
		rem echo Historico já existe
	) ELSE (
		rem echo Criando Histórico
		md Historico
	)



:planilha
	IF EXIST "Historico/%FORMATO%.xlsm" (
		rem echo Planilha já Existe
	) ELSE (	
		rem echo Criando Planilha
		copy "00 Branco.xlsm" "Historico\%FORMATO%.xlsm"
	)	


:exe
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
	rem Configurações
	attrib +s +h "%caminho%\*.exe"
	attrib +s +h "%caminho%\*.txt"
	attrib +s +h "%caminho%\*.bat"
	attrib -s -h "%caminho%\*.xlsm"


:abrir
	rem Abrir Planilha
	call "%caminho%\%FORMATO%.exe"

:sair
	rem Sair da Aplicação
	Exit

:aberto
	cls
	echo Planilha já está aberta

:menu
	rem Menu
	Title Abrindo Planilha...
	goto historico
	goto planilha
	goto exe
	goto abrir
	goto conf