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
	pb := progressbar("Initializing browser...","w400")

	loop 3
	{
		eventlog("Browser open attempt " A_index)
		pb.set(33*A_Index)
		; wb := wbOpen()																	; start/activate an Chrome/Edge instance
		; if IsObject(wb) {
		; 	break
		wb := ""
	}
	if !IsObject(wb) {
		eventlog("Failed to open browser.")
		pb.close
		MsgBox("Failed to open browser","PrevGrab error",262160)
		ExitApp
	}


;#endregion

ExitApp

;#region == GUI Elements ===============================================================
class progressbar
{
	__New(params*) {
		param := ""
		title := ""
		subtitle := ""

		for val in params {
			if (param="") && (val~="([wW]\d+).*?([hH]\d+)?") {							; matches "w000" or "h000"
				param := val
			} 
			else if (title="") {														; first non-param text
				title := val
			} 
			else if (subtitle="") {														; second non-param text
				subtitle := val
			}
		}

		par := parseParam(param)

		pb := Gui()
		pb.Opt("+Border +AlwaysOnTop -SysMenu -Caption")
		if (title) {
			pb.SetFont("s16")
			pb.AddText(par.W " Center vTitle",title)
		}
		pb.AddProgress(par.W " " par.H " " par.C " vPercent")
		if (subtitle) {
			pb.SetFont("s12")
			pb.AddText(par.W " Center vSubtitle",subtitle)
		}
		this.gui := pb
		pb.Show()

		parseParam(param) {
			width := (RegExMatch(" " param " ","\W[wW](\d+)\W",&par)) ? par[0] : "w200"
			height := (RegExMatch(" " param " ","\W[hH]\w+\W",&par)) ? par[0] : "h12"
			color := (RegExMatch(" " param " ","\W[cC]\w+\W",&par)) ? par[0] : "cBlue"
			return {W:width,H:height,C:color}
		}
	}

	set(val) {
		try this.gui["Percent"].Value := val
	}
	title(val) {
		try this.gui["Title"].Value := val
	}
	sub(val) {
		try this.gui["Subtitle"].Value := val
	}
	close() {
		try {
			this.gui.Destroy()
			this.gui := ""
		}
	}
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