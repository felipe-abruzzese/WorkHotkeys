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




run, C:\Users\felipe.ramalho\Documents\Macros\AUTOCAD.ahk



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
	run C:\Users\felipe.ramalho\Desktop\2019 - BANCO DE HORAS FELIPE.xlsx
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


;^y:: ; testes legais
;	send, ^{c}
;	FileAppend, %clipboard%, C:\Users\felipe.ramalho\Documents\Macros\Test.txt

;return



^q:: ; testes legais
	keywait, ctrl
	keywait, q
	Blockinput, on
	sendraw, R6.
	blockinput, off
	return
	
	/*
	mousemove, 1484, 532
	Loop, 30{
	click, WD
	}
	inputbox, numero, Tipo de Ponto, Digite o tipo de ponto (1 2 ou 3)
	inputbox, sufixo, Sufixo, Digite o Sufixo
	if (numero == 1){
	click, 1505, 517
	
	send, {end}
	send, +{home}
	send, R2.
	send, %sufixo%

	}if (numero == 2){
	click, 1508, 534 
	send, {end}
	send, +{home}
	send, R2.
	send, %sufixo%
	
	}if (numero == 3){
	click, 1477, 409
	send, {end}
	send, +{home}
	send, R2.
	send, %sufixo%
	
	}

	blockinput, off
return
*/


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

;**************************************************************** COMANDOS APONTAMENTO ATIVIDADES *************************************************************************************************************

#IfWinActive ahk_class XLMAIN
^#+k:: ; PREENCHIMENTO AUTOMÁTICO DE GTI
;setkeydelay, 500
;setwindelay, 500
;setcontroldelay, 500

t = 250
tlong = 1500

keywait, ctrl
keywait, Lwin
keywait, Lshift
keywait, k

blockinput, on
	
	Send {f2}
	Send {end}
	Send +{home}
	Send ^{c}
	Send {esc}
	sleep, %t%
	winactivate ahk_exe chrome.exe
	winwaitactive ahk_exe chrome.exe
		sendinput ^{v}
		sleep, %t%
		sendinput {tab}
		sleep, %t%
	winactivate ahk_class XLMAIN
	winwaitactive ahk_class XLMAIN
		sendinput {right 2}
		sendinput {up}
		sendinput ^{c}
		sleep, %t%
	winactivate ahk_exe chrome.exe
	winwaitactive ahk_exe chrome.exe		
		sendinput +{end}
		sendinput ^{v}
		sendinput {tab}
		sleep, %t%
	winactivate ahk_class XLMAIN
	winwaitactive ahk_class XLMAIN
		sendinput {right 4}
		sendinput ^{c}
		sleep, %t%
	winactivate ahk_exe chrome.exe
	winwaitactive ahk_exe chrome.exe		
		sendinput +{end}
		sendinput ^{v}
		sendinput {tab}
		sleep, %t%
	winactivate ahk_class XLMAIN
	winwaitactive ahk_class XLMAIN
		sendinput {left 4}
		sendinput {down}
		sendinput ^{c}
		sleep, %t%
	winactivate ahk_exe chrome.exe
	winwaitactive ahk_exe chrome.exe		
		sendinput +{end}
		sendinput ^{v}
		sendinput {tab 2}
		sleep, %t%
	winactivate ahk_class XLMAIN
	winwaitactive ahk_class XLMAIN
		sendinput {left}
		sendinput ^{down}
		sendinput {left}
		Send {f2}
		Send {end}
		Send +{home}
		Send ^{c}
		Send {esc}
		sleep, %t%
	
	send {tab}
	send ^{up}
	send {left}
	send {down 3}
	loop, %clipboard% {
				
			sendinput {right}
			sendinput ^{c}
			sendinput {right}
			sleep, %t%
		winactivate ahk_exe chrome.exe
		winwaitactive ahk_exe chrome.exe		
			sendinput %clipboard%
			sendinput {tab}
			sleep, %t%
		winactivate ahk_class XLMAIN
		winwaitactive ahk_class XLMAIN
			Send {f2}
			Send {end}
			Send +{home}
			Send ^{c}
			Send {esc}
			sendinput {right}
			sleep, %t%
		winactivate ahk_exe chrome.exe
		winwaitactive ahk_exe chrome.exe		
			sleep, %t% ;%t%*4
			sendinput %clipboard% ;UNIORG
			sleep, %tlong% ;%t%*4
			sendinput {enter}
			sendinput {tab}
			sleep, %t%
			
			loop, 3 {
			winactivate ahk_class XLMAIN
			winwaitactive ahk_class XLMAIN
				Send ^{c}
				sendinput {right}
				sleep, %t%
			winactivate ahk_exe chrome.exe
			winwaitactive ahk_exe chrome.exe		
				sendinput %clipboard%
				sendinput {tab}
				sleep, %t%
				}
		winactivate ahk_class XLMAIN
		winwaitactive ahk_class XLMAIN
			Send {f2}
			Send {end}
			Send +{home}
			Send ^{c}
			Send {esc}
			sendinput {right}
			sleep, %t%	
		winactivate ahk_exe chrome.exe
		winwaitactive ahk_exe chrome.exe		
			sendinput +{end}
			sendinput %clipboard%
			sendinput {tab}
			sleep, %t%
		winactivate ahk_class XLMAIN
		winwaitactive ahk_class XLMAIN
			Send ^{c}
			sleep, %t%
		winactivate ahk_exe chrome.exe
		winwaitactive ahk_exe chrome.exe		
			sendinput %clipboard%
			sendinput {tab}
			sleep, %t%
		winactivate ahk_class XLMAIN
		winwaitactive ahk_class XLMAIN	
			sendinput {down}
			sendinput ^{left}
			sleep, %t%
			}
	sendinput {down}
	sendinput ^{left}
	sendinput {down}
blockinput, off
return
	


