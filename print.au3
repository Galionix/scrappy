 #include <WindowsConstants.au3>
 #include <SendMessage.au3>


Send("!{TAB}")
Sleep(200)
Send("^v")
 _SendMessage(ControlGetHandle("", "", ""), $WM_PASTE)
Send("{ENTER}")
Sleep(200)
Send("!{TAB}")