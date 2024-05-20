/*	PrevGrab v2

*/

#Requires AutoHotkey v2
#SingleInstance Force
#Include "includes"
SendMode("Input")
SetWorkingDir A_ScriptDir
SetTitleMatchMode("2")

;#region == CONFIGURATION ==============================================================
	gl := {}

	if InStr(A_ScriptDir,"AhkProjects") {
		gl.isDevt := true
	} else {
		gl.isDevt := false
	}
	; A_Args[1] := "ftp"				;*******************************

	gl.TRRIQ_path := A_ScriptDir
	gl.files_dir := gl.TRRIQ_path "\files"
	gl.pdfTemp := gl.TRRIQ_path "\pdfTemp"
	wq := xml.new(gl.TRRIQ_path "\worklist.xml")
	
	gl.settings := readIni("settings")
	
	gl.enroll_ct := 0
	gl.inv_ct := 0
	gl.t0 := A_TickCount

;#endregion

;#region == MAIN LOOP ==================================================================
	eventlog("Initializing.")
	pb := progressbar("Initializing webdriver...")

	

;#endregion

ExitApp

;#region == GUI Elements ===============================================================
progressbar(title:="",param:="",title2:="",title3:="") {
/*	Creates a minimal progress bar using Title, Params, Subtitle
	If first var is a pbar object, param=percentage, title2=new title, title3=subtitle
*/
	local width:="w200", height:="h12", color:="cBlue"
	if IsObject(title) {
		if IsNumber(param) {
			title["Percent"].Value := param
		}
		if (title2) {
			try {
				title["Title"].Value := title2
			} 
		}
		if (title3) {
			try {
				title["Subtitle"].Value := title3
			} 
		}
		return
	}
	width := (RegExMatch(" " param " ","\W[wW](\d+)\W",&par)) 
		? par[0] : "w200"
	height := (RegExMatch(" " param " ","\W[hH]\w+\W",&par))
		? par[0] : "h12"
	color := (RegExMatch(" " param " ","\W[cC]\w+\W",&par))
		? par[0] : "cBlue"
	pbar := Gui()
	pbar.Opt("+Border +AlwaysOnTop -SysMenu -Caption")
	if (title) {
		pbar.SetFont("s16")
		pbar.AddText(width " Center vTitle",title)
	}
	pbar.AddProgress(width " " height " " color " vPercent")
	if (title2) {
		pbar.SetFont("s12")
		pbar.AddText(width " Center vSubtitle",title2)
	}
	pbar.Show()
	return pbar
}

;#endregion

;#region == TEXT Elements ==============================================================
eventlog(event,verbosity:=1) {
	/*	verbose 1 or 0 from ini
		verbosity default 1
		verbosity set 0 if only during verbose
	*/
		global gl
		
		score := verbosity + gl.settings["verbose"]
		if (score<1) {
			return
		}
		user := A_UserName
		comp := A_ComputerName
		sessDate := FormatTime(A_Now,"yyyy.MM")											; FormatTime, sessdate, A_Now, yyyy.MM
		now := FormatTime(A_Now,"yyyy.MM.dd||HH:mm:ss") 								; FormatTime, now, A_Now, yyyy.MM.dd||HH:mm:ss
		name := gl.TRRIQ_path "\logs\" . sessdate . ".log"
		txt := now " [" user "/" comp "] PREVGRAB: " event "`n"
		filePrepend(txt,name)
}

FilePrepend( Text, Filename ) { 
/*	from haichen http://www.autohotkey.com/board/topic/80342-fileprependa-insert-text-at-begin-of-file-ansi-text/?p=510640
*/
	file:= FileOpen(Filename, "rw")
	text .= File.Read()
	file.pos:=0
	File.Write(text)
	File.Close()
}

readIni(section) {
/*	Reads a set of variables

	[section]					==	 		var1 := res1, var2 := res2
	var1=res1
	var2=res2
	
	[array]						==			array := ["ccc","bbb","aaa"]
	=ccc
	=bbb
	=aaa
	
	[objet]						==	 		objet := {aaa:10,bbb:27,ccc:31}
	aaa:10
	bbb:27
	ccc:31
*/
	global
	local x, i, key, val, k, v
		, i_res
		, i_type := []
		, i_lines := []
		, iniFile := ".\files\prevgrab.ini"
	i_type.var := i_type.obj := i_type.arr := false

	x:=IniRead(iniFile,section)
	loop parse x, "`n", "`r"
	{
		i := A_LoopField
		if (i~="(?<!`")[=]") 															; find = not preceded by "
		{
			if (i ~= "^=") {															; starts with "=" is an array list
				i_type.arr := true
				i_res := Array()
			} else {																	; "aaa=123" is a var declaration
				i_type.var := true
			}
		} 
		else																			; does not contain a quoted =
		{
			if (i~="(?<!`")[:]") {														; find : not preceded by " is an object
				i_type.obj := true
				i_res := Map()
		} else {																		; contains neither = nor : can be an array list
				i_type.arr := true
				i_res := Array()
			}
		}
	}
	if ((i_type.obj) + (i_type.arr) + (i_type.var)) > 1 {								; too many types, return error
		return error
	}
	Loop parse x, "`n","`r"																; now loop through lines
	{
		i := A_LoopField
		if (i_type.var) {
			key := strX(i,"",1,0,"=",1,1)
			val := trim(strX(i,"=",1,1,"",1,0),'`"')
			k := &key
			v := &val
			%k% := %v%
		}
		if (i_type.obj) {
			key := trim(strX(i,"",1,0,":",1,1),'`"')
			val := trim(strX(i,":",1,1,"",0),'`"')
			i_res[key] := val
		}
		if (i_type.arr) {
			i := RegExReplace(i,"^=")													; remove preceding =
			i_res.push(trim(i,'`"'))
		}
	}
	return i_res
}

;#endregion

#Include xml2.ahk
#Include strx2.ahk