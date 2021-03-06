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

run C:\Program Files (x86)\Evernote\Evernote\Evernote.exe
run https://todoist.com/
Return



^!#p:: ; Abrir pasta Projetos
keywait, ctrl
keywait, Lwin
keywait, alt
Blockinput, on
ifwinexist PROJETOS
{
winactivate PROJETOS
}
else
{
run \\Servidor\projetos\PROJETOS
winactivate PROJETOS
}
Blockinput, off
Return

^#s:: ;ABRIR WINDOWS SPY
keywait, ctrl
keywait, Lwin
Blockinput, on
run C:\Program Files\AutoHotkey\AU3_Spy.exe
Blockinput, off
Return


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;



^#r:: ; REINICIA A MACRO INICIAR TRABALHOS
Run C:\Users\abruz\Desktop\One Drive\OneDrive\Felipe\Softwares\Macros\notebook.ahk
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

^#n::
IfWinExist, ahk_class Notepad++
WinActivate
Run C:\Program Files\Notepad++\notepad++.exe
Return



;**************************************COMANDOS EXCEL*****************************************************

#IfWinActive ahk_class XLMAIN
;::Houv:: / Houveram quedas pontuais de disponibilidade prejudicando o todo que seria xx% / Nenhuma a��o � necess�ria
;:*:CC::=CONCATENAR("SMS,";
;:*:SS::;",SIGNAL")
;:*:ii::=SEERRO(SE(E(�NDICE(
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
send {f2}
Send {end}
send +{home}
Send ^{c}
send {esc}
}
Return


;***************************************************************************************ENVIO DE OR�AMENTO AUTOM�TICO *****************************************************************************************************

^#k:: ;Pedido autom�tico de or�amento
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
Sendraw, Pedido de Or�amento

;Come�a o Texto do Email:
Send, {tab}
sendraw, �
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
sendraw, Segue abaixo os itens a serem or�ados.
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
Sendraw, Qualquer d�vida estou a disposi��o
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





;*******************************************************************************************************TEMPOR�RIO**********************************************************************************************************


