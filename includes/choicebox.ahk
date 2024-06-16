#Requires AutoHotkey v2

/*	choiceBox - MsgBox-like dialog with multiple choice buttons
	Returns button text, or "xClose" if [x] button
	icon "I" = Info
	icon "Q" = Question
	icon "E" = Error
	icon "!" = Exclamation
	vert = completely vertical alignment
	img = filename of image to place below textbox
 */
choiceBox(title:="",text:="",buttons:=[],icon:="",vert:="v",img:="") {
	cMsg := Gui()
	hwnd := cMsg.Hwnd
	cMsg.Opt("+ToolWindow +AlwaysOnTop")

	MyIcon := ( icon = "I" ) 															; INFO
		? "icon5" 	
		: (icon = "Q") 																	; QUESTION
			? "icon3" 	
			: (icon = "E") 																; ERROR
				? "icon4" 	
				: (icon = "!")															; EXCLAMATION 
					? "icon2"
					: ""
	cMsg.AddPicture(MyIcon,A_WinDir "\system32\user32.dll")
	cMsg.AddText("x+12 yp w180 r8 section",text)
	cMsg.Title := title

	if (img) {
		; Cmsg.AddPicture(
		; 	(A_Index=1)
		; 		? "x+12 ys Section "
		; 		: "xs y+3 Section ", 
		; 	(IsObject(img[A_Index]) ? img[A_Index] : "")
		; )
	}

	for lbl in buttons
	{
		cMsg.AddButton(
			((vert="v")
				? ""
				: "xs+90 ys " )
			. ((lbl~="^\*")
				? "Default "
				: " " )
			. "w150 "  
			, RegExReplace(lbl,"^\*")
		)
		.OnEvent("Click",cMsgButton)
	}

	cMsg.Show("AutoSize")
	cMsg.OnEvent("Close",cMsgClose)
	res := ""

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
