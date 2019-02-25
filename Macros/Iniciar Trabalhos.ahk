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


^#n:: ; Enumerar automaticamente

keywait, ctrl
keywait, Lwin
Blockinput, on
InputBox, Inicio, Valor Inicial, Digite o Valor Inicial, 
InputBox, Fim, Valor Final, Digite o valor mais alto.

while Inicio <= Fim {

send, {Tab 2}
send, %Inicio%
++Inicio
send, {Tab 2}
send, r
}
Blockinput, off
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
Run C:\Users\felipe.ramalho\Documents\Macros\Iniciar Trabalhos.ahk
Return


^!0::Run calc.exe ; RODAR CALCULADORA
Return


^#'::
IfWinexist, ahk_class XLMAIN
WinActivate
{
Send {F2}
Send {home}
Send {'}
Send {enter}
}
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
;::Houv:: / Houveram quedas pontuais de disponibilidade prejudicando o todo que seria xx% / Nenhuma ação é necessária
;:*:CC::=CONCATENAR("SMS,";
;:*:SS::;",SIGNAL")
;:*:ii::=SEERRO(SE(E(ÍNDICE(
;:*:RR::CORRESP(
^q::
{
Send, {alt}
Send, {c}
Send, {s}
Send, {i}
}
^+c::
{
Send {f2}
Send {end}
Send +{home}
Send ^{c}
Send {esc}
}
Return


;***************************************************************************************ENVIO DE ORÇAMENTO AUTOMÁTICO *****************************************************************************************************

^#k:: ;Pedido automático de orçamento
#WinActivateForce
#IfWinnotActive ahk_class XLMAIN
WinActivate
keywait, ctrl
keywait, Lwin
keywait, k

blockinput, on

setwindelay, 250
setkeydelay, 250

send, ^{pgup 10}
send, ^{pgdn 2}

Send, ^{left 3}
Send, ^{up 3}

Send, {down 16}
Send, {right 2}

Send, ^{c}

sleep, 500

#Ifwinexist ahk_exe OUTLOOK.EXE
WinActivate ahk_exe OUTLOOK.EXE
winwaitactive, ahk_exe OUTLOOK.EXE

sleep, 500

Send, {alt}
Send, {c}
Send, {t}
Send, {o}
Sleep, 500
Send, ^{v}
Send, {tab 2}
Sendraw, Pedido de Orçamento

;Começa o Texto do Email:
Send, {tab}
sendraw, À
send {space}
WinActivate, ahk_class XLMAIN
winwaitactive, ahk_class XLMAIN
Send, {up 4}
Send, ^{c}
WinActivate, ahk_exe OUTLOOK.EXE
winwaitactive, ahk_exe OUTLOOK.EXE
Send, {alt}
Send, {m}
Send, {v}
Send, {t}

send, {enter 2}
sendraw, Segue abaixo os itens a serem orçados.
send, {enter 3}
sleep, 500

WinActivate, ahk_class XLMAIN
winwaitactive, ahk_class XLMAIN
send, {down 9}
sleep, 500
send, ^+{down}
Send, +{right 2}
Send, ^{c}
sleep, 500
WinActivate, ahk_exe OUTLOOK.EXE
winwaitactive, ahk_exe OUTLOOK.EXE
sleep, 500
Send, {alt}
Send, {m}
Send, {v}
Send, {f}

send, {enter 2}
Sendraw, Qualquer dúvida estou a disposição
send, {enter 2}
Sendraw, Att.
Send, {enter 2}

Blockinput, off
Return

;*************************************************************************************** Auto Preenchimento Autocad *****************************************************************************************************

; Macro auto preenchimento

^#!k::
#Ifwinexist ahk_exe acad.exe
WinActivate ahk_exe acad.exe
winwaitactive, ahk_exe acad.exe
sendraw, find
send, {enter}
sleep, 2000
sendraw, yy
send, {tab 3}
send, {f}
sleep, 1000
send, {enter}
sleep, 1000
send, {tab 2}
WinActivate, ahk_exe notepad.exe
winwaitactive, ahk_exe notepad.exe
send ^{home}
loop, 5
{
send, +{end}
send, ^{c}
WinActivate ahk_exe acad.exe
winwaitactive, ahk_exe acad.exe
send, ^{v}
send, {tab 2}
send, {r}
send, {tab 2}
WinActivate, ahk_exe notepad.exe
winwaitactive, ahk_exe notepad.exe
send {down}
}
return





;********************************************************************************************************EVERNOTE**********************************************************************************************************





;*******************************************************************************************************TEMPORÁRIO**********************************************************************************************************


