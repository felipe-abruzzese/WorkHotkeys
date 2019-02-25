;
; AutoHotkey Version: 1.x
; Language:       English
; Platform:       Win9x/NT
; Author:         A.N.Other <myemail@nowhere.com>
;
; Script Function:
;	Template script (you can customize this template by editing "ShellNew\Template.ahk" in your Windows folder)
;

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; - CABEÇALHO - ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

ifwinexist ahk_exe acad.exe
	{
		run, C:\Users\felipe.ramalho\Documents\Macros\AUTOCAD.ahk
	}


;AUTORELOAD
	sleep, 20000
	reload
	
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; - INICIO - ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

^#!i:: ; Iniciar Trabalhos

	;run C:\Program Files (x86)\Evernote\Evernote\Evernote.exe
	run C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2013\Outlook 2013
	run https://todoist.com/
	run https://www.kemp-gti.com.br/index.asp
	run C:\Users\felipe.ramalho\Desktop\q-dir.exe
Return

^+v:: ; COLAR COMO TEXTO NÃO FORMATADO

;OUTLOOK
	IfWinActive ahk_exe OUTLOOK.EXE
	{
		keywait, ctrl
		keywait, alt
		Blockinput, on
		Sendinput, {lalt}
		Send, {m}
		send, {v}
		send, {t}
		blockinput, off
	}
	else

;WORD
	ifwinactive ahk_exe WINWORD.EXE
	{
		keywait, ctrl
		keywait, alt
		Blockinput, on
		Sendinput, {lalt}
		Send, {c}
		send, {v}
		send, {t}
		blockinput, off
	}
	else
;EXCEL
	ifwinactive ahk_exe EXCEL.EXE
	{
		keywait, ctrl
		keywait, alt
		Blockinput, on
		Sendinput, {lalt}
		Send, {c}
		send, {v}
		send, {v}
		blockinput, off
	}
	else
;AUTOCAD
	ifwinactive ahk_exe acad.exe
	{
		keywait, ctrl
		keywait, alt
		Blockinput, on
		sendraw, pasteblock
		send {enter}
		blockinput, off
	}
	Return


#y:: ; testes legais
	send %clipboard%
return



^#s:: ;ABRIR WINDOWS SPY
	keywait, ctrl
	keywait, Lwin
	Blockinput, on
	run C:\Users\felipe.ramalho\Documents\ahk\WindowSpy.ahk
	Blockinput, off
Return


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

^#i:: ; INICIA O EVERNOTE
	ifwinexist, ahk_class ENMainFrame
		winactivate
	else
		Run C:\Users\felipe.ramalho\AppData\Local\Apps\Evernote\Evernote\Evernote.exe
return

^#r:: ; REINICIA A MACRO INICIAR TRABALHOS
	reload
Return


^!0::Run calc.exe ; RODAR CALCULADORA
Return


#IfWinActive, ahk_exe q-dir.exe
{
return::
keywait, return
blockinput, on
send {enter}
sleep, 100
send {tab}
sleep, 100
send, +{tab}
blockinput, off
Return
}


;******************************************************************************************* COMANDOS EXCEL *************************************************************************************************************

#IfWinActive ahk_class XLMAIN
^+c:: ; Copiar o Conteúdo da Célula
{
Send {f2}
Send {end}
Send +{home}
Send ^{c}
Send {esc}
}
Return



	
