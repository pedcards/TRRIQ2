#Requires AutoHotkey v2

/*	choiceBox - MsgBox-like dialog with multiple choice buttons
		"*Button" will be default button
		Returns button text, or "xClose" if [x] button

	Options:
		+iconI = Info ('Q'=Question, 'E','X'=Error, '!'=Exclamation)
		+v(ert) = completely vertical alignment
		+tw___ = textbox width (default 240)
		+bw___ = button width (default 150)
		+slim = button height to fit
		+fat = button height extra padding
		
		img = filename of image to replace icon, separated by comma
 */
choiceBox(title:="",text:="",buttons:=[],opts*) {

	textW := 240 , btnW := "w150 "
	vert := img := res := ""
	thisIcon := "icon5"
	rows := "r2 "

	for val in opts
	{
		if RegExMatch(val " ","i)\+icon(.)\W",&i) {
			thisIcon := (i[1] = "I" ) ? "icon5" 										; INFO
					: (i[1] = "Q") ? "icon3"											; QUESTION
					: (i[1] ~= "E|X") ? "icon4"											; ERROR
					: (i[1] = "!") ? "icon2"											; EXCLAMATION 
					: "icon5"
		}
		if (val~="i)\+v(ert)?") {
			vert := true
		}
		if RegExMatch(val " ","i)\+tw(\d{3,})",&w) {
			textW := w[1]
		}
		if RegExMatch(val " ","i)\+bw(\d{3,})",&w) {
			btnW := "w" w[1] " "
		}
		if (val~="i)\+slim") {
			rows := ""
		}
		if (val~="i)\+fat") {
			rows := "r3 "
		}
		if FileExist(val) {
			img := val
		}
	}

	cMsg := Gui()
	hwnd := cMsg.Hwnd
	cMsg.Opt("+ToolWindow +AlwaysOnTop")

	cMsg.Title := title
	if (img) {
		cMsg.AddPicture( , img)
	} else {
		cMsg.AddPicture(thisIcon,A_WinDir "\system32\user32.dll")
	}
	tbox := cMsg.AddText("x+12 yp w" textW " Section",text "`n")

	cMsg.AddText((vert) ? "" : "ys-16")													; Set position for buttons
	for lbl in buttons
	{
		cMsg.AddButton(rows
			. ((lbl~="^\*") ? "Default " : " ")
			. btnW
			, RegExReplace(lbl,"^\*")
		)
		.OnEvent("Click",cMsgButton)
	}

	cMsg.Show("AutoSize")
	cMsg.OnEvent("Close",cMsgClose)

	WinWaitClose("ahk_id " hwnd)
	return res

	cMsgButton(var,*) {
		res := var.text
		try cMsg.Destroy()
	}
	cMsgClose(Cmsg) {
		res := "xClose"
		try Cmsg.Destroy()
	}
}
