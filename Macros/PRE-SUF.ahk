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

;**************************************************************** COMANDOS APONTAMENTO ATIVIDADES *************************************************************************************************************

^#V::

	loop, 5{
	send, {f2}
	send, {home}
	sendraw, BR-SAO-ASJ-ELE-PB-
	send, {end}
	sendraw, -R0
	send, {tab}
	}
	return
	