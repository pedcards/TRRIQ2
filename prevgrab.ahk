/*	PrevGrab v2

*/

#Requires AutoHotkey v2
#SingleInstance Force
#Include "includes"
SendMode("Input")
SetWorkingDir A_ScriptDir
SetTitleMatchMode("2")

__Config:
{
	gl := {}
	pb := progressbar("Initializing...")
	progressbar(pb,90)

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

}

ExitApp

__GUI_elements:
progressbar(title:="",param:="w200 cBlue",title2:="",title3:="") {
/*	Creates a minimal progress bar using Title and Params
	If first var is a pbar object, set the percentage
*/
	if IsObject(title) {
		title["Percent"].Value := param
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
		pbar.AddText(width " Center",title)
	}
	pbar.AddProgress(width " " height " " color " vPercent")
	pbar.Show()
	return pbar
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

#Include xml2.ahk
#Include strx2.ahk