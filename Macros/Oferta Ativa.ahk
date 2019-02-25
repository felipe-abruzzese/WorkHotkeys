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