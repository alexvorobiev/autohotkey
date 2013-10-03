;#IfWinActive emacs  ; if in emacs
+Capslock::Capslock ; make shift+Caps-Lock the Caps Lock toggle
Capslock::Control   ; make Caps Lock the control button
;#IfWinActive        ; end if in emacs

#`::Run c:\emacs\bin\emacsclientw.exe -n -e "(make-capture-frame)"


^+F10::
  ClipSave := clipboard 
  Clipboard =           ; empty clipboard
  Send ^c               ; copy selection
  ClipWait
  sel := clipboard    ; save to variable
  Clipboard =           ; empty clipboard
  WinGetActiveTitle, Title ; get window title
  url :=
  TitleClean :=
  
  if WinActive("ahk_class MozillaWindowClass") 
    or WinActive("ahk_class Chrome_WidgetWin_0")
    or WinActive("ahk_class Chrome_WidgetWin_1") { 
    Send ^l               ; select
    Sleep 25
    Send ^c
    ClipWait
    url := clipboard
    TitleClean := RegExReplace(Title," - Google Chrome")
  
  } else if WinActive("ahk_class XLMAIN") {
    xl := ComObjActive("Excel.Application")
    url := xl.ActiveWorkbook.FullName
    TitleClean := RegExReplace(Title,"(.*) - ")  ; remove Microsoft Excel... 

  }
  ;MsgBox % "[[" . url . "][" . TitleClean . "]] sel"

  Clipboard := "[[" . url . "][" . TitleClean . "]] sel"
  Run c:\emacs\bin\emacsclientw.exe -n -e "(make-capture-frame)"

return
