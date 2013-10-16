; Remap CapsLock in Emacs
;
#IfWinActive ahk_class Emacs 
{
  +Capslock::Capslock ; make shift+Caps-Lock the Caps Lock toggle
  Capslock::Control   ; make Caps Lock the control button
}

#IfWinActive        ; end if in emacs

; Win-` creates application-depenedent org-mode link on the clipboard
; and creates emacs frame with org-capture:
;
;  Chrome: [[url][page title]] selection
;  Excel:  [[path][file name]] selection
;
#`::
  ClipSave := clipboard 
  Clipboard =           ; empty clipboard
  Send ^c               ; copy selection
  ClipWait, 2
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
    ClipWait, 2
    url := clipboard
    TitleClean := RegExReplace(Title," - Google Chrome")
  
  } else if WinActive("ahk_class XLMAIN") {
    xl := ComObjActive("Excel.Application")
    url := xl.ActiveWorkbook.FullName
    TitleClean := RegExReplace(Title,"(.*) - ")  ; remove Microsoft Excel... 
  }

  Clipboard := "[[" . url . "][" . TitleClean . "]]" . sel
  Run c:\emacs\bin\emacsclientw.exe -n -e "(make-capture-frame)"

return
