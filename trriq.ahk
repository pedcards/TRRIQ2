/*	TRRIQ v2

*/

#Requires AutoHotkey v2
#SingleInstance Force
#Include "includes"
SendMode("Input")
SetWorkingDir A_ScriptDir
SetTitleMatchMode("2")

;#region == CONFIGURATION ==============================================================
	pb := progressbar("w400","TRRIQ initializing..."," ")
	
	/*	User and path
	*/
	gl := {}
	gl.user := A_UserName
	gl.userinstance := substr(tobase(A_TickCount,36),-4)
	gl.comp := A_ComputerName
	gl.wqfileDT := FileGetTime(".\files\wqupdate")
	gl.runningVer := FileGetTime(A_ScriptName)

	/*	Determine PROD, DEVT, TEST
	*/
	SplitPath(A_ScriptDir,,&fileDir)
	if InStr(fileDir, "AhkProjects") {
		gl.isDevt := true
		try FileDelete(".lock")
		path:=readIni("devtpaths")
		eventlog(">>>>> Started in DEVT mode.")
	} else {
		gl.isDevt := false
		path:=readIni("paths")
		eventlog(">>>>> Started in PROD mode. " A_ScriptName " ver " substr(gl.runningVer,1,12) " " A_Args[1])
	}
	if InStr(fileDir,"TEST") {
		gl.isDevt := True
		eventlog("***** launched from TEST folder.")
	}
	if ObjHasOwnProp(A_Args,"launch") {
		eventlog("***** launched from legacy shortcut.")
		FileAppend(A_Now ", " gl.user "|" gl.userinstance "|" A_ComputerName "`n", ".\files\legacy.txt")
		MsgBox(
			"Obsolete TRRIQ shortcut!`n`n"
			. "Please notify Igor Gurvits or Jim Gray to update the shortcut on this machine: " A_ComputerName
			, "Shortcut error"
			, 0x30
		)
	}

	/*	Read ini vars
	*/
	readini("setup")

	/*	Get location info
	*/
	gl.wksVoid := StrSplit(gl.wksVM, "|")
	pb.title("Identifying workstation...")
	wks := GetLocation()
	if !(wksLoc := wks.location) {
		pb.Destroy
		MsgBox("No clinic location specified!`n`nExiting","Location error",262160)
		ExitApp
	}

	sites := wks.getSites(wksLoc)
	; sites.tracked	(aka sites)						= sites we are tracking
	; sites.ignored (aka sites0)					= sites we are not tracking <tracked>N</tracked> in wkslocation
	; sites.long									= {CIS:TAB}
	; sites.code									= {"MAIN":7343} 4 digit code for sending facility
	; sites.facility								= {"MAIN":"GB-SCH-SEATTLE"}

	/*	Read outdocs.csv for Cardiologist and Fellow names 
	*/
	Docs := readDocs()																	; returns Docs[site][idx].name

	/*	Generate worklist.xml if missing
	*/
	gl.wq_filename := path.data "worklist.xml"
	if fileexist(gl.wq_filename) {
		wq := XML(gl.wq_filename)
	} else {
		wq := XML("<root/>")
		wq.addElement("/root","pending")
		wq.addElement("/root","done")
		wq.save(gl.wq_filename)
	}

	/*	Read call schedule (Electronic Forecast and Qgenda)
	*/
	fcVals := readIni("Forecast")
	updateCall()

	/*	Initialize rest of vars and strings
	*/
	pb.title("Initializing variables")
	pb.sub("Demographics")
	pb.set()
	demVals := readIni("demVals")																		; valid field names for parseClip()

	pb.sub("Indication codes")
	indCodes := readIni("indCodes")																		; valid indications
	; *** try to use indCodes array by itself
	; ***
	; for key,val in indCodes																				; in option string indOpts
	; {
	; 	tmpVal := strX(val,"",1,0,":",1)
	; 	tmpStr := strX(val,":",1,1,"",0)
	; 	indOpts .= tmpStr "|"
	; }

	pb.sub("Monitor strings")
	monStrings := readIni("Monitors")																	; Monitor key strings
	monOrderType := {}
	monSerialStrings := {}
	monPdfStrings := {}
	monEpicEAP := {}
	for key,val in monStrings
	{
		; Monitor letter code "H": Order abbrev "HOL": Order list dur "24-hr": Regex type "Pr|Hol": Regex S/N "Mortara": Epic EAP "CVCAR102:HOLTER MONITOR 24 HOUR" 
		el := strSplit(val,":")
		monOrderType[el.2]:=el.3																		; String matches for order <mon>
		monSerialStrings[el.2]:=el.5																	; Regex matches for S/N strings
		monPdfStrings[el.1]:=el.2																		; Abbrev based on PDF fname
		monEpicEAP[el.2]:=el.6																			; Epic EAP codes for monitors
	}

	pb.sub("HL7 map")
	; initHL7()																							; HL7 definitions
	hl7DirMap := {}

	pb.sub("Reading EP list")
	epList := readIni("epRead")																			; reading EP
	for key in epList																					; option string epStr
	{
		epStr .= key "|"
	}
	epStr := Trim(epStr,"|")

	pb.sub("Save recent Cygnus logs")
	; saveCygnusLogs("all")
		
;#endregion

ExitApp

;#region == GUI elements ===============================================================

;#endregion

;#region == TEXT functions ==============================================================
eventlog(event) {
	global gl

	sessDate := FormatTime(A_Now,"yyyy.MM")											; FormatTime, sessdate, A_Now, yyyy.MM
	now := FormatTime(A_Now,"yyyy.MM.dd||HH:mm:ss") 								; FormatTime, now, A_Now, yyyy.MM.dd||HH:mm:ss
	fname := ".\logs\" . sessdate . ".log"
	txt := now " [" gl.user "/" gl.comp "/" gl.userinstance "] " event "`n"
	filePrepend(txt,fname)
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
	local x, i, key, val
		, i_res
		, i_type := []
		, i_lines := []
		, iniFile := ".\files\trriq.ini"
	i_type.var := i_type.obj := i_type.arr := false

	x:=IniRead(iniFile,section)
	loop parse x, "`n", "`r"
	{
		i := A_LoopField
		if (i ~= "^=") {																; starts is "=" is an array list
			i_type.arr := true
			i_res := Array()
		} 
		else if (i ~= "(?<!`")[:]") {													; ":" not preceded by " is object
			i_type.obj := true
			i_res := Map()
		} 
		else if (i~="(?<!`")[=]") { 													; "aaa=123" is a var declaration
			i_type.var := true
			i_res := ""
		}
		else {																			; contains neither = nor : can be an array list
			i_type.arr := true
			i_res := Array()
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
			gl.%key% := val
			; k := &key
			; v := &val
			; %k% := %v%
		}
		if (i_type.obj) {
			key := trim(strX(i,"",1,0,":",1,1),'`"')
			val := trim(strX(i,":",1,1,"",0),'`"')
			i_res.%key% := val
		}
		if (i_type.arr) {
			i := RegExReplace(i,"^=")													; remove preceding =
			i_res.push(trim(i,'`"'))
		}
	}
	return i_res
}

ParseDate(x) {
	mo := ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
	moStr := "Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec"
	dSep := "[ \-\._/]"
	date := {yyyy:"",mmm:"",mm:"",dd:"",date:""}
	time := {hr:"",min:"",sec:"",days:"",ampm:"",time:""}

	x := RegExReplace(x,"[,\(\)]")

	if (x~="\d{4}.\d{2}.\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?Z") {
		x := RegExReplace(x,"[TZ]","|")
	}
	if (x~="\d{4}.\d{2}.\d{2}T\d{2,}") {
		x := RegExReplace(x,"T","|")
	}

	if RegExMatch(x,"i)(\d{1,2})" dSep "(" moStr ")" dSep "(\d{4}|\d{2})",&d) {			; 03-Jan-2015
		date.dd := zdigit(d[1])
		date.mmm := d[2]
		date.mm := zdigit(objhasvalue(mo,d[2]))
		date.yyyy := d[3]
		date.date := trim(d[0])
	}
	else if RegExMatch(x,"\b(\d{4})[\-\.](\d{2})[\-\.](\d{2})\b",&d) {					; 2015-01-03
		date.yyyy := d[1]
		date.mm := zdigit(d[2])
		date.mmm := mo[d[2]]
		date.dd := zdigit(d[3])
		date.date := trim(d[0])
	}
	else if RegExMatch(x,"i)(" moStr "|\d{1,2})" dSep "(\d{1,2})" dSep "(\d{4}|\d{2})",&d) {	; Jan-03-2015, 01-03-2015
		date.dd := zdigit(d[2])
		date.mmm := objhasvalue(mo,d[1]) 
			? d[1]
			: mo[d[1]]
		date.mm := objhasvalue(mo,d[1])
			? zdigit(objhasvalue(mo,d[1]))
			: zdigit(d[1])
		date.yyyy := (d[3]~="\d{4}")
			? d[3]
			: (d[3]>50)
				? "19" d[3]
				: "20" d[3]
		date.date := trim(d[0])
	}
	else if RegExMatch(x,"i)(" moStr ")\s+(\d{1,2}),?\s+(\d{4})",&d) {					; Dec 21, 2018
		date.mmm := d[1]
		date.mm := zdigit(objhasvalue(mo,d[1]))
		date.dd := zdigit(d[2])
		date.yyyy := d[3]
		date.date := trim(d[0])
	}
	else if RegExMatch(x,"\b(19\d{2}|20\d{2})(\d{2})(\d{2})((\d{2})(\d{2})(\d{2})?)?\b",&d)  {	; 20150103174307 or 20150103
		date.yyyy := d[1]
		date.mm := d[2]
		date.mmm := mo[d[2]]
		date.dd := d[3]
		if (d[1]) {
			date.date := d[1] "-" d[2] "-" d[3]
		}
		
		time.hr := d[5]
		time.min := d[6]
		time.sec := d[7]
		if (d[5]) {
			time.time := d[5] ":" d[6] . strQ(d[7],":###")
		}
	}

	if RegExMatch(x,"i)(\d+):(\d{2})(:\d{2})?(:\d{2})?(.*)?(AM|PM)?",&t) {				; 17:42 PM
		hasDays := (t[4]) ? true : false 											; 4 nums has days
		time.days := (hasDays) ? t[1] : ""
		time.hr := trim(t[1+hasDays])
		time.min := trim(t[2+hasDays]," :")
		time.sec := trim(t[3+hasDays]," :")
		if (time.min>59) {
			time.hr := floor(time.min/60)
			time.min := zDigit(Mod(time.min,60))
		}
		if (time.hr>23) {
			time.days := floor(time.hr/24)
			time.hr := zDigit(Mod(time.hr,24))
			DHM:=true
		}
		time.ampm := trim(t[5])
		time.time := trim(t[0])
	}

	return {yyyy:date.yyyy, mm:date.mm, mmm:date.mmm, dd:date.dd, date:date.date
			, YMD:date.yyyy date.mm date.dd
			, YMDHMS:date.yyyy date.mm date.dd zDigit(time.hr) zDigit(time.min) zDigit(time.sec)
			, MDY:date.mm "/" date.dd "/" date.yyyy
			, MMDD:date.mm "/" date.dd 
			, hrmin:zdigit(time.hr) ":" zdigit(time.min)
			, days:zdigit(time.days)
			, hr:zdigit(time.hr), min:zdigit(time.min), sec:zdigit(time.sec)
			, ampm:time.ampm, time:time.time
			, DHM:zdigit(time.days) ":" zdigit(time.hr) ":" zdigit(time.min) " (DD:HH:MM)" 
			, DT:date.mm "/" date.dd "/" date.yyyy " at " zdigit(time.hr) ":" zdigit(time.min) ":" zdigit(time.sec) }
}

zDigit(x) {
; Returns 2 digit number with leading 0
	return SubStr("00" x, -2)
}

strQ(var1,txt,null:="") {
/*	Print Query - Returns text based on presence of var
	var1	= var to query
	txt		= text to return with ### on spot to insert var1 if present
	null	= text to return if var1="", defaults to ""
*/
	return (var1="") ? null : RegExReplace(txt,"###",var1)
}

;#endregion

;#region == OTHER FUNCTIONS ============================================================
ObjHasValue(aObj, aValue, rx:="") {
	for key, val in aObj
		if (rx="RX") {																	; argument 3 is "RX" 
			if (aValue="") {															; null aValue in "RX" is error
				return false
			}
			if (val ~= aValue) {														; val=text, aValue=RX
				return key
			}
			if (aValue ~= val) {														; aValue=text, val=RX
				return key
			}
		} else {
			if (val = aValue) {															; otherwise just string match
				return key
			}
		}
	return false																		; fails match, return err
}

ToBase(n,b) {
/*	from https://autohotkey.com/board/topic/15951-base-10-to-base-36-conversion/
	n >= 0, 1 < b <= 36
*/
	Return (n < b ? "" : ToBase(n//b,b)) . ((d:=mod(n,b)) < 10 ? d : Chr(d+55))
}
	
filecheck() {
	if FileExist(".lock") {
		err:=0
		pb := progressbar("Waiting to clear lock","File write queued...")
		loop 50 {
			if (FileExist(".lock")) {
				pb.set(p)
				Sleep 100
				p += 2
			} else {
				err:=1
				break
			}
		}
		if !(err) {
			pb.close
			return error
		}
		pb.close
	} 
	return
}

;#endregion

#Include xml2.ahk
#Include strx2.ahk
#Include progressbar.ahk
#Include HostName.ahk
#Include updateData.ahk

#Include Peep.v2.ahk																	; This is only for debugging