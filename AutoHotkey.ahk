;#IfWinActive emacs  ; if in emacs
+Capslock::Capslock ; make shift+Caps-Lock the Caps Lock toggle
Capslock::Control   ; make Caps Lock the control button
;#IfWinActive        ; end if in emacs

#`::Run c:\emacs\bin\emacsclientw.exe -n -e "(make-capture-frame)"


^+F10::
  if WinActive("ahk_class MozillaWindowClass") 
  or WinActive("ahk_class Chrome_WidgetWin_0")
  or WinActive("ahk_class Chrome_WidgetWin_1") { 
    ClipSave := clipboard 
    Clipboard =           ; empty clipboard
    Send ^c               ; copy selection
    ClipWait
    sel := clipboard    ; save to variable
    Clipboard =           ; empty clipboard
    Send ^l               ; select
    Sleep 25
    Send ^c
    ClipWait
    url := clipboard
    WinGetActiveTitle, Title
    TitleClean := RegExReplace(Title," - Google Chrome")
    Clipboard := ClipSave
    MsgBox % "[[" . url . "][" . TitleClean . "]] sel"
  }
return

^+F9::
xl := ComObjActive("Excel.Application")
path := xl.ActiveWorkbook.Path
MsgBox % path
return
