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

;Resumo: 	^#i ---> Tem Interesse
; 			^#c ---> Não tem Interesse
;			^#n ---> Não Atende
;			^#t ---> Telefone Não existe
;			^#r ---> Telefone errado



^#i:: ; Tem Interesse

ifwinexist houseCrm - Plano & Plano - Google Chrome
winactivate, houseCrm - Plano & Plano - Google Chrome
winwaitactive, houseCrm - Plano & Plano - Google Chrome
keywait, lwin
keywait, ctrl
blockinput, on
send, {home}
sleep, 100
loop, 3
{
click, 215,288 ; NOME
}
loop, 3
{
sleep, 100
send, ^{c}
}
winactivate, OFERTA ATIVA.xlsx - Excel
winwaitactive, OFERTA ATIVA.xlsx - Excel
loop, 5
{
send, ^{up}
sleep, 100
send, ^{left}
}
send, {right}
sleep, 100
send, ^{down}
sleep, 100
send, {down}
sleep, 100
send, {f2}
sleep, 100
send, ^{v}
sleep, 100
send, {tab}
sleep, 100
winactivate, houseCrm - Plano & Plano - Google Chrome
winwaitactive, houseCrm - Plano & Plano - Google Chrome
loop, 3
{
click, 648,276 ; TELEFONE
}
loop, 3
{
sleep, 100
send, ^{c}
}
winactivate, OFERTA ATIVA.xlsx - Excel
winwaitactive, OFERTA ATIVA.xlsx - Excel
send, {f2}
sleep, 100
send, ^{v}
sleep, 100
send, {tab}
sleep, 100
winactivate, houseCrm - Plano & Plano - Google Chrome
winwaitactive, houseCrm - Plano & Plano - Google Chrome
loop, 3
{
click, 164,354 ; EMAIL
}
loop, 3
{
sleep, 100
send, ^{c}
}
winactivate, OFERTA ATIVA.xlsx - Excel
winwaitactive, OFERTA ATIVA.xlsx - Excel
send, {f2}
sleep, 100
send, ^{v}
sleep, 100
send, {tab}
sleep, 100
winactivate, houseCrm - Plano & Plano - Google Chrome
winwaitactive, houseCrm - Plano & Plano - Google Chrome
loop, 3
{
click, 163,206 ; CAMPANHA
}
loop, 3
{
sleep, 100
send, ^{c}
}
winactivate, OFERTA ATIVA.xlsx - Excel
winwaitactive, OFERTA ATIVA.xlsx - Excel
send, {f2}
sleep, 100
send, ^{v}
sleep, 100
send, {tab}
sleep, 100
winactivate, houseCrm - Plano & Plano - Google Chrome
winwaitactive, houseCrm - Plano & Plano - Google Chrome
click, 385,651 ; OBS
send, ^{home}
sleep, 100
send, ^+{end}
loop, 3
{
sleep, 100
send, ^{c}
}
winactivate, OFERTA ATIVA.xlsx - Excel
winwaitactive, OFERTA ATIVA.xlsx - Excel
send, {f2}
sleep, 100
send, ^{v}
sleep, 100
send, {tab}
sleep, 100
sendraw, TEM_INTERESSE
send, {enter}
winactivate, houseCrm - Plano & Plano - Google Chrome
winwaitactive, houseCrm - Plano & Plano - Google Chrome
click, 79, 453 ; Caixa - tem interesse
sleep, 100
click, 707,523 ; salvar e próximo cliente
blockinput, off
;msgbox, Comando encerrado
return


^#c:: ; Não tem interesse

ifwinexist houseCrm - Plano & Plano - Google Chrome
winactivate, houseCrm - Plano & Plano - Google Chrome
winwaitactive, houseCrm - Plano & Plano - Google Chrome
keywait, lwin
keywait, ctrl
blockinput, on
send, {home}
sleep, 100
loop, 3
{
click, 215,288 ; NOME
}
loop, 3
{
sleep, 100
send, ^{c}
}
winactivate, OFERTA ATIVA.xlsx - Excel
winwaitactive, OFERTA ATIVA.xlsx - Excel
loop, 5
{
send, ^{up}
sleep, 100
send, ^{left}
}
send, {right}
sleep, 100
send, ^{down}
sleep, 100
send, {down}
sleep, 100
send, {f2}
sleep, 100
send, ^{v}
sleep, 100
send, {tab}
sleep, 100
winactivate, houseCrm - Plano & Plano - Google Chrome
winwaitactive, houseCrm - Plano & Plano - Google Chrome
loop, 3
{
click, 648,276 ; TELEFONE
}
loop, 3
{
sleep, 100
send, ^{c}
}
winactivate, OFERTA ATIVA.xlsx - Excel
winwaitactive, OFERTA ATIVA.xlsx - Excel
send, {f2}
sleep, 100
send, ^{v}
sleep, 100
send, {tab}
sleep, 100
winactivate, houseCrm - Plano & Plano - Google Chrome
winwaitactive, houseCrm - Plano & Plano - Google Chrome
loop, 3
{
click, 164,354 ; EMAIL
}
loop, 3
{
sleep, 100
send, ^{c}
}
winactivate, OFERTA ATIVA.xlsx - Excel
winwaitactive, OFERTA ATIVA.xlsx - Excel
send, {f2}
sleep, 100
send, ^{v}
sleep, 100
send, {tab}
sleep, 100
winactivate, houseCrm - Plano & Plano - Google Chrome
winwaitactive, houseCrm - Plano & Plano - Google Chrome
loop, 3
{
click, 163,206 ; CAMPANHA
}
loop, 3
{
sleep, 100
send, ^{c}
}
winactivate, OFERTA ATIVA.xlsx - Excel
winwaitactive, OFERTA ATIVA.xlsx - Excel
send, {f2}
sleep, 100
send, ^{v}
sleep, 100
send, {tab}
sleep, 100
winactivate, houseCrm - Plano & Plano - Google Chrome
winwaitactive, houseCrm - Plano & Plano - Google Chrome
click, 385,651 ; OBS
send, ^{home}
sleep, 100
send, ^+{end}
loop, 3
{
sleep, 100
send, ^{c}
}
winactivate, OFERTA ATIVA.xlsx - Excel
winwaitactive, OFERTA ATIVA.xlsx - Excel
send, {f2}
sleep, 100
send, ^{v}
sleep, 100
send, {tab}
sleep, 100
sendraw, NAO_TEM_INTERESSE
sleep,100
send, {enter}
winactivate, houseCrm - Plano & Plano - Google Chrome
winwaitactive, houseCrm - Plano & Plano - Google Chrome
click, 98, 477 ; Caixa - descartar
sleep, 100
click, 201,492 ; Click Caixa 
sleep, 100
send, {down 5}
sleep, 100
send, {enter}
sleep, 100
click, 707,523 ; salvar e próximo cliente
blockinput, off

return


^#n:: ; Não atende

ifwinexist houseCrm - Plano & Plano - Google Chrome
winactivate, houseCrm - Plano & Plano - Google Chrome
winwaitactive, houseCrm - Plano & Plano - Google Chrome
keywait, lwin
keywait, ctrl
blockinput, on
send, {home}
sleep, 100
loop, 3
{
click, 215,288 ; NOME
}
loop, 3
{
sleep, 100
send, ^{c}
}
winactivate, OFERTA ATIVA.xlsx - Excel
winwaitactive, OFERTA ATIVA.xlsx - Excel
loop, 5
{
send, ^{up}
sleep, 100
send, ^{left}
}
send, {right}
sleep, 100
send, ^{down}
sleep, 100
send, {down}
sleep, 100
send, {f2}
sleep, 100
send, ^{v}
sleep, 100
send, {tab}
sleep, 100
winactivate, houseCrm - Plano & Plano - Google Chrome
winwaitactive, houseCrm - Plano & Plano - Google Chrome
loop, 3
{
click, 648,276 ; TELEFONE
}
loop, 3
{
sleep, 100
send, ^{c}
}
winactivate, OFERTA ATIVA.xlsx - Excel
winwaitactive, OFERTA ATIVA.xlsx - Excel
send, {f2}
sleep, 100
send, ^{v}
sleep, 100
send, {tab}
sleep, 100
winactivate, houseCrm - Plano & Plano - Google Chrome
winwaitactive, houseCrm - Plano & Plano - Google Chrome
loop, 3
{
click, 164,354 ; EMAIL
}
loop, 3
{
sleep, 100
send, ^{c}
}
winactivate, OFERTA ATIVA.xlsx - Excel
winwaitactive, OFERTA ATIVA.xlsx - Excel
send, {f2}
sleep, 100
send, ^{v}
sleep, 100
send, {tab}
sleep, 100
winactivate, houseCrm - Plano & Plano - Google Chrome
winwaitactive, houseCrm - Plano & Plano - Google Chrome
loop, 3
{
click, 163,206 ; CAMPANHA
}
loop, 3
{
sleep, 100
send, ^{c}
}
winactivate, OFERTA ATIVA.xlsx - Excel
winwaitactive, OFERTA ATIVA.xlsx - Excel
send, {f2}
sleep, 100
send, ^{v}
sleep, 100
send, {tab}
sleep, 100
winactivate, houseCrm - Plano & Plano - Google Chrome
winwaitactive, houseCrm - Plano & Plano - Google Chrome
click, 385,651 ; OBS
send, ^{home}
sleep, 100
send, ^+{end}
loop, 3
{
sleep, 100
send, ^{c}
}
winactivate, OFERTA ATIVA.xlsx - Excel
winwaitactive, OFERTA ATIVA.xlsx - Excel
send, {f2}
sleep, 100
send, ^{v}
sleep, 100
send, {tab}
sleep, 100
sendraw, NAO_ATENDE
sleep,100
send, {enter}
winactivate, houseCrm - Plano & Plano - Google Chrome
winwaitactive, houseCrm - Plano & Plano - Google Chrome
click, 87, 521 ; Caixa - descartar
sleep, 100
click, 190,523 ; Click Caixa 
sleep, 100
send, {down 2}
sleep, 100
send, {enter}
sleep, 100
click, 707,523 ; salvar e próximo cliente
blockinput, off

return

^#t:: ; Telefone Não existe

ifwinexist houseCrm - Plano & Plano - Google Chrome
winactivate, houseCrm - Plano & Plano - Google Chrome
winwaitactive, houseCrm - Plano & Plano - Google Chrome
keywait, lwin
keywait, ctrl
blockinput, on
send, {home}
sleep, 100
click, 98, 477 ; Caixa - descartar
sleep, 100
click, 201,492 ; Click Caixa 
sleep, 100
send, {down 7} ; Telefone Não existe
sleep, 100
send, {enter}
sleep, 100
click, 707,523 ; salvar e próximo cliente
blockinput, off

return

^#r:: ; Telefone errado

ifwinexist houseCrm - Plano & Plano - Google Chrome
winactivate, houseCrm - Plano & Plano - Google Chrome
winwaitactive, houseCrm - Plano & Plano - Google Chrome
keywait, lwin
keywait, ctrl
blockinput, on
send, {home}
sleep, 100
click, 98, 477 ; Caixa - descartar
sleep, 100
click, 201,492 ; Click Caixa 
sleep, 100
send, {down 6} ; Telefone errado
sleep, 100
send, {enter}
sleep, 100
click, 707,523 ; salvar e próximo cliente
blockinput, off

return

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