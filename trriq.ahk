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
	; monOrderType := {}
	; monSerialStrings := {}
	; monPdfStrings := {}
	; monEpicEAP := {}

	monTypes := Map()
	for key,val in monStrings
	{
		monTypes[key] := Map() 
		el := strSplit(val,":")
		monTypes[key]["type"] := el[1]																	; Monitor letter code "H"
		monTypes[key]["abbrev"] := el[2]																; Abbrev for PDF fname "HOL"
		monTypes[key]["duration"] := el[3]																; Order list dur "24-hr"
		monTypes[key]["modelRegex"] := el[4]															; Mon type regex "Pr|Hol"
		monTypes[key]["serial"] := el[5]																; S/n regex "Mortara|Mini SL"
		monTypes[key]["EAP"] := el[6]																	; Epic EAP "CVCAR102^HOLTER MONITOR 24 HOUR"

		; monOrderType[el.2]:=el.3																		; String matches for order <mon>
		; monSerialStrings[el.2]:=el.5																	; Regex matches for S/N strings
		; monPdfStrings[el.1]:=el.2																		; Abbrev based on PDF fname
		; monEpicEAP[el.2]:=el.6																			; Epic EAP codes for monitors
	}

	pb.sub("HL7 map")
	hl7 := getHL7()																						; HL7 definitions
	hl7DirMap := {}

	pb.sub("Reading EP list")
	epList := readIni("epRead")																			; reading EP
	; for key in epList																					; option string epStr
	; {
	; 	epStr .= key "|"
	; }
	; epStr := Trim(epStr,"|")

	pb.sub("Screen dimensions")
	dims := getDims()

	pb.sub("Save recent Cygnus logs")
	saveCygnusLogs("all")

	pb.title("Cleaning old .bak files")
	pb.sub("")
	cleanBakFiles()
	pb.close()

		
;#endregion

;#region == MAIN LOOP ==================================================================
	Loop
	{
		PhaseGUI()
		WinWaitClose("TRRIQ Dashboard")
		
		; if (phase="HolterUpload") {
		; 	eventlog("Start Holter Connect.")
		; 	hcPhase := "Transfer"
		; 	HolterConnect(hcPhase)
		; }
	}
	
	; saveCygnusLogs("all")
	; checkPreventiceOrdersOut()
	; cleanDone()

	ExitApp
	
;#endregion

;#region == GUI elements ===============================================================
getDims() {
	res := Map()

	res.screen := Map()
	res.screen.DPI := A_ScreenDPI
	res.screen.H := A_ScreenHeight
	res.screen.W := A_ScreenWidth

	res.phase := Map()
	res.phase.lvH := 450
	res.phase.lvW := 720

	return res
}

PhaseGUI() {
	global dims
	
	phase := Gui()
	phase.Opt("+AlwaysOnTop")


/*
	Gui, Add, Text, % "x" lvW+40 " y15 w200 vPhaseNumbers", "`n`n"
	Gui, Add, GroupBox, % "x" lvW+20 " y0 w220 h65"
	
	Gui, Font, Bold
	Gui, Add, Button
		, Y+10 wp h40 gPhaseRefresh
		, Refresh lists
	Gui, Add, Button
		, Y+10 wp h40 gPrevGrab Disabled
		, Check Preventice inventory
	Gui, Add, Text, wp h50																; space between top buttons and lower buttons
	Gui, Add, Text, Y+10 wp h24 Center, Register/Prepare a`nHOLTER or EVENT MONITOR
	Gui, Add, Button
		, Y+10 wp h40 vRegister gPhaseOrder DISABLED
		, No active orders
	Gui, Add, Text, wp h30
	Gui, Add, Text, Y+10 wp Center		, Transmit
	Gui, Add, Text, Y+1 wp Center H100	, BG MINI
	Gui, Font, Normal

	GuiControlGet, btn2, Pos, BG MINI

	btnW := 79
	btnH := 61

	Gui, Add, Picture
	, % "Y" btn2Y+20 " X" btn2X+70
	. " w" btnW " h" btnH " "
	. " +0x1000 vHolterUpload gPhaseTask"
	, .\files\BGMini.png

	tmpsite := RegExReplace(sites,"TRI\|")
	tmpsite := wksloc="Main Campus" ? tmpsite : RegExReplace(tmpsite,site.tab "\|",site.tab "||")
	Gui, Add, Tab3																		; add Tab bar with tracked sites
		, -Wrap x10 y10 w%lvW% h%lvH% vWQtab +HwndWQtab
		, % "ORDERS|" 
		. (wksloc="Main Campus" ? "INBOX||" : "") 
		. "Unread|ALL|" tmpsite
	GuiControlGet, wqDim, Pos, WQtab
	lvDim := "W" wqDimW-25 " H" wqDimH-35
	
	if (wksloc="Main Campus") {
		GuiControl 
			, Enable
			, Check Preventice inventory

		Gui, Tab, INBOX
		Gui, Add, Listview
			, % "-Multi Grid BackgroundSilver " lvDim " greadWQlv vWQlv_in hwndHLV_in"
			, filename|Name|MRN|DOB|Location|Study Date|wqid|Type|Need FTP
		Gui, ListView, WQlv_in
		LV_ModifyCol(1,"0")																; filename and path, "0" = hidden
		LV_ModifyCol(2,"160")															; name
		LV_ModifyCol(3,"60")															; mrn
		LV_ModifyCol(4,"80")															; dob
		LV_ModifyCol(5,"80")															; site
		LV_ModifyCol(6,"80")															; date
		LV_ModifyCol(7,"2")																; wqid
		LV_ModifyCol(8,"40")															; ftype
		LV_ModifyCol(9,"70 Center")														; ftp
		CLV_in := new LV_Colors(HLV_in,true,false)
		CLV_in.Critical := 100
	}
	
	Gui, Tab, ORDERS
	Gui, Add, Listview
		, % "-Multi Grid BackgroundSilver ColorRed " lvDim " greadWQorder vWQlv_orders hwndHLV_orders"
		, filename|Order Date|Name|MRN|Ordering Provider|Monitor
	Gui, ListView, WQlv_orders
	LV_ModifyCol(1,"0")																	; filename and path (hidden)
	LV_ModifyCol(2,"80")																; date
	LV_ModifyCol(3,"140")																; Name
	LV_ModifyCol(4,"60")																; MRN
	LV_ModifyCol(5,"100")																; Prov
	LV_ModifyCol(6,"70")																; Type
	
	Gui, Tab, Unread
	Gui, Add, Listview
		, % "-Multi Grid BackgroundSilver ColorRed " lvDim " vWQlv_unread hwndHLV_unread"
		, Name|MRN|Study Date|Processed|Monitor|Ordering|Assigned EP
	Gui, ListView, WQlv_unread
	LV_ModifyCol(1,"140")																; Name
	LV_ModifyCol(2,"60")																; MRN
	LV_ModifyCol(3,"80")																; Date
	LV_ModifyCol(4,"80")																; Processed
	LV_ModifyCol(5,"70")																; Mon Type
	LV_ModifyCol(6,"80")																; Ordering
	LV_ModifyCol(7,"80")																; Assigned EP

	Gui, Tab, ALL
	Gui, Add, Listview
		, % "-Multi Grid BackgroundSilver " lvDim " gWQtask vWQlv_all hwndHLV_all"
		, ID|Enrolled|FedEx|Uploaded|Notes|MRN|Enrolled Name|Device|Provider|Site
	Gui, ListView, WQlv_all
	LV_ModifyCol(1,"0")																	; wqid (hidden)
	LV_ModifyCol(2,"60")																; date
	LV_ModifyCol(3,"40 Center")															; FedEx
	LV_ModifyCol(4,"60")																; uploaded
	LV_ModifyCol(5,"40 Center")															; Notes
	LV_ModifyCol(6,"60")																; MRN
	LV_ModifyCol(7,"140")																; Name
	LV_ModifyCol(8,"130")																; Ser Num
	LV_ModifyCol(9,"100")																; Prov
	LV_ModifyCol(10,"80")																; Site
	CLV_all := new LV_Colors(HLV_all,true,false)
	CLV_all.Critical := 100
	
	Loop, parse, sites, |
	{
		i := A_Index
		site := A_LoopField
		Gui, Tab, % site
		Gui, Add, Listview
			, % "-Multi Grid BackgroundSilver " lvDim " gWQtask vWQlv"i " hwndHLV"i
			, ID|Enrolled|FedEx|Uploaded|Notes|MRN|Enrolled Name|Device|Provider
		Gui, ListView, WQlv%i%
		LV_ModifyCol(1,"0")																	; wqid (hidden)
		LV_ModifyCol(2,"60")																; date
		LV_ModifyCol(3,"40 Center")															; FedEx
		LV_ModifyCol(4,"60")																; uploaded
		LV_ModifyCol(5,"40 Center")															; Notes
		LV_ModifyCol(6,"60")																; MRN
		LV_ModifyCol(7,"140")																; Name
		LV_ModifyCol(8,"130")																; Ser Num
		LV_ModifyCol(9,"100")																; Prov
		CLV_%i% := new LV_Colors(HLV%i%,true,false)
		CLV_%i%.Critical := 100
	}
	WQlist()
	
	Menu, menuSys, Add, Change clinic location, changeLoc
	Menu, menuSys, Add, Generate late returns report, lateReport
	Menu, menuSys, Add, Generate registration locations report, regReport
	Menu, menuSys, Add, Update call schedules, updateCall
	Menu, menuSys, Add, CheckMWU, checkMWUapp											; position for test menu
	Menu, menuHelp, Add, About TRRIQ, menuTrriq
	Menu, menuHelp, Add, Instructions..., menuInstr
	Menu, menuAdmin, Add, Toggle admin mode, toggleAdmin
	Menu, menuAdmin, Add, Clean tempfiles, CleanTempFiles
	Menu, menuAdmin, Add, Send notification email, sendEmail
	Menu, menuAdmin, Add, Find pending leftovers, cleanPending
	Menu, menuAdmin, Add, Fix WQ device durations, fixDuration							; position for test menu
	Menu, menuAdmin, Add, Recover DONE record, recoverDone
	Menu, menuAdmin, Add, Check running users/versions, runningUsers
	Menu, menuAdmin, Add, Create test order, makeEpicORM
		
	Menu, menuBar, Add, System, :menuSys
	if (user~="i)tchun1|docte") {
		Menu, menuBar, Add, Admin, :menuAdmin
	}
	Menu, menuBar, Add, Help, :menuHelp
	
	Gui, Menu, menuBar
	Gui, Show,, TRRIQ Dashboard

	if (adminMode) {
		Gui, Color, Fuchsia
		Gui, Show,, TRRIQ Dashboard - ADMIN MODE
	}
	
	SetTimer, idleTimer, 500
	return
*/
}

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

;#region == PREVENTICE FUNCTIONS =======================================================
saveCygnusLogs(all:="") {
/*	Save copy of Cygnus logs per machine per user
	"all" creates new mirror
	Toggle in trriq.ini
*/
	folder := A_AppData "\Cygnus\Logs"
	logpath := ".\logs\Cygnus\" A_ComputerName "\" A_UserName
	today := "Log_" A_YYYY "-" A_MM "-" A_DD ".log"
	if (all) {																			; any value will copy entire folder 
		DirCopy(folder, logpath "\", 1)													; trailing \ copies folder to logpath
	} else {
		FileCopy(folder "\" today, logpath "\" today, 1)
	}
	FileSetTime( , ".\logs\Cygnus\" A_ComputerName "\" A_UserName)
	FileSetTime( , ".\logs\Cygnus\" A_ComputerName)
	
	Return
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

cleanBakFiles() {
	Loop files ".\bak\*.bak"
	{
		count++
	}
	Loop files ".\bak\*.bak"
	{
		pb.set(100*A_Index/count)
		dt := dateDiff(A_Now,RegExReplace(A_LoopFileName,"\.bak"),"Days")
		if (dt > 7) {
			FileDelete(".\bak\" A_LoopFileName)
		}
	}
}

;#endregion

#Include xml2.ahk
#Include strx2.ahk
#Include progressbar.ahk
#Include HostName.ahk
#Include updateData.ahk
#Include hl7.ahk

#Include Peep.v2.ahk																	; This is only for debugging