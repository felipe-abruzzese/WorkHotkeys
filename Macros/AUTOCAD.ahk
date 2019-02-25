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


^!v:: ; Deletar Layout em massa
;Autocad
IfWinActive ahk_exe acad.exe
{
keywait, ctrl
keywait, alt
Blockinput, on
Sendraw, Layout
Send, {d}
winactivate ahk_exe notepad.exe
winwaitactive ahk_exe notepad.exe
Send, ^{home}
send, +{end}
send, ^{c}
send, {del 2}
winactivate ahk_exe acad.exe
winwaitactive ahk_exe acad.exe
send, {clipboard}
blockinput, off
}
Return