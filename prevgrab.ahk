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
	gl.inv_ct := gl.inv_tot := 0
	gl.t0 := A_TickCount
	gl.clip := gl.clip0 := ""
	gl.FAIL := 0
	gl.prevtxt := ""

;#endregion

;#region == MAIN LOOP ==================================================================
	eventlog("Initializing.")
	pb := progressbar("Initializing browser...","w400"," ")

	loop 3
	{
		eventlog("Browser open attempt " A_index)
		pb.set(33*A_Index)
		wb := wbOpen()																	; start/activate an Chrome/Edge instance
		if IsObject(wb) {
			break
		}
	}
	if !IsObject(wb) {
		eventlog("Failed to open browser.")
		pb.close
		MsgBox("Failed to open browser","PrevGrab error",262160)
		ExitApp
	} else {
		pb.set(100)
	}
	webStr := {}
	wb.visible := gl.settings.isVisible													; for progress bars
	wb.capabilities.HeadlessMode := gl.settings.isHeadless								; for Chrome/Edge window
	wb.capabilities.IncognitoMode := gl.settings.isIncognito							; incognito mode does not save passwords
	gl.Page := wb.NewSession()															; Session in gl.Page

	if ObjHasOwnProp(A_Args,"ftp") {
		webStr.FTP := readIni("str_ftp")
		gl.login := readIni("str_ftpLogin")

		PreventiceWebGrab("ftp")
		gl.FAIL := gl.wbFail
	} else {
		; webStr.Enrollment := readIni("str_Enrollment")
		webStr.Inventory := readIni("str_Inventory")
		gl.login := readIni("str_Login")

		; PreventiceWebGrab("Enrollment")
		PreventiceWebGrab("Inventory")
		if (gl.inv_ct < gl.inv_tot) {
			gl.FAIL := true
		}
		FileDelete(gl.files_dir "\prev.txt*")											; writeout each one regardless
		FileAppend(gl.prevtxt, gl.files_dir "\prev.txt")
		eventlog("Enroll " gl.enroll_ct ", Inventory " gl.inv_ct ". (" round((A_TickCount-gl.t0)/1000,2) " sec)")
	
	}
	pb.close

	eventlog("Closing webdriver.")
	gl.Page.Close()
	wb.Exit()

	if (gl.FAIL) {																		; Note when a table had failed to load
		MsgBox("Downloads failed.","PrevGrab",262160)
		eventlog("Critical hit: Downloads failed.")
	} else {
		MsgBox("Preventice update complete!","PrevGrab",262160)
	}
	
	ExitApp

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
		return

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

;#region == WEB Elements ===============================================================
wbOpen() {
/*	Use Rufaydium class https://github.com/Xeo786/Rufaydium-Webdriver
	to use Google Chrome or Microsoft Edge webdriver to retrieve webpage
*/
	try cr32Ver := FileGetVersion("C:\Program Files (x86)\Google\Chrome\Application\chrome.exe")
	try cr64Ver := FileGetVersion("C:\Program Files\Google\Chrome\Application\chrome.exe")
	try mseVer := FileGetVersion("C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe")
	
	if (cr64Ver) {
		verNum := cr64Ver
		driver := "chromedriver"
		eventlog("Found Chrome (x64) version " verNum)
	} Else
	if (cr32Ver) {
		verNum := cr32Ver
		driver := "chromedriver"
		eventlog("Found Chrome (x86) version " verNum)
	} Else
	if (mseVer) {
		verNum := mseVer
		driver := "msedgedriver"
		eventlog("Found Edge (x86) version " verNum)
	} Else {
		eventlog("Could not find installed Chrome or Edge.")
		Return
	}
	Num :=  strX(verNum,"",0,1,".",1,1)

	exe := A_ScriptDir "\bin\" driver ".exe"
	if !FileExist(exe) {
		eventlog("Could not find matching driver. Attempt download.")
	}
	wb := Rufaydium(driver,,0,0)

	return wb
}

PreventiceWebGrab(phase) {
	web := webStr.%phase%
	pb.title("Checking " phase)

	wbUrl(web.url)																		; load URL, return DOM in gl.Page
	if (gl.wbFail) {
		return
	}
	wbWaitBusy(gl.settings.webwait)
	prvFunc := web.fx
	loop
	{
		pb.title("Parsing inventory")
		pb.sub("Page " A_Index)
		try WinHide("Save password ahk_exe chrome.exe")

		tbl := gl.Page.getElementsByClassName(web.tbl)[0]								; get the Main Table
		if !IsObject(tbl) {
			eventlog("*** " phase " *** No matching table.")
			gl.FAIL := true
			return
		}
		
		body := tbl.getElementsByClassName(web.tblBody)[0]
		
		done := %prvFunc%(body)		; parsePreventiceEnrollment() or parsePreventiceInventory() or parsePreventiceFTP()
		
		if (done=0) {																	; no new records returned
			break
		}
		
		PreventiceWebPager(phase,web.changed,web.btn)
	}
	
	return

}

wbUrl(url) {
/*	Open a URL
*/
	gl.wbFail := true
	loop 3																				; Number of attempts to permit redirects
	{
		try 
		{
			pb.sub("Launching URL attempt " A_Index)
			eventlog("Navigating to " url " (attempt " A_index ").")
			gl.Page.Navigate(url) 														; load URL
			if !(wbWaitBusy(gl.settings.webwait)) {										; msec before fails
				eventlog("Failed to load.")
			}
			if !(gl.Page.URL = url) {
				eventlog("Redirected.",0)
				sleep 1000
			}
		}
		catch as e
		{
			eventlog("wbUrl failed with msg: " stregX(e.message "`n","",1,0,"[\r\n]+",1))
			if instr(e.message,"The RPC server is unavailable") {
				eventlog("Reloading DOM...")
				wb := wbOpen()
			}
			continue
		}
		
		if instr(gl.Page.URL,gl.login.string) {
			pb.sub("Sending login")
			loginErr := preventiceLogin()
			eventlog("Login " ((loginErr) ? "submitted." : "attempted."))
			try WinHide("Save password ahk_exe chrome.exe")
			sleep 1000
		}
		if (gl.Page.URL=url) {
			eventlog("Succeeded.",0)
			gl.wbFail := false
			return
		}
		else {
			eventlog("Landed on " gl.Page.URL)
			sleep 500
		}
	}
	gl.wbFail := true
	eventlog("Failed all attempts " url)
	return
}

wbWaitBusy(maxTick) {
	startTick:=A_TickCount
	
	while InStr(gl.Page.html,"table-body ng-hide") {									; class="table-body ng-hide" present while rendering FTP list
		if (A_TickCount-startTick > maxTick) {
			eventlog(gl.Page.url " timed out.")
			return false
		}
		sleep 500
	} 
	while InStr(gl.Page.html,"{{progress}}") {											; present when rendering FTP login page
		if (A_TickCount-startTick > maxTick) {
			eventlog(gl.Page.url " timed out.")
			return false
		}
		sleep 500
	} 
	; while gl.Page.IsLoading() {															; wait until done loading
	; 	if (A_TickCount-startTick > maxTick) {
	; 		eventlog(gl.Page.url " timed out.")
	; 		return false																; break loop if time exceeds maxTick
	; 	}
	; 	checkBtn("Message from webpage","OK")											; check if err window present and click OK button
	; 	sleep 200
	; }
	return A_TickCount-startTick
}

checkBtn(txt,btn) {
	if (errHWND:=WinExist(txt)) {
		ControlClick(%btn%,"ahk_id " %errHWND%)
		eventlog("Message dialog clicked '" btn "'.")
		sleep 200
	}
	return
}

preventiceLogin() {
/*	Need to populate and submit user login form
*/
	gl.Page
		.getElementById(gl.login.attr_user)
		.SendKey(gl.login.user_name)													; .value := gl.login.user_name
	
	gl.Page
		.getElementById(gl.login.attr_pass)
		.SendKey(gl.login.user_pass)													; .value := gl.login.user_pass
	
	gl.Page
		.getElementByID(gl.login.attr_btn)
		.click()
	
	if !(wbWaitBusy(gl.settings.webwait)) {												; wait until done loading
		return false
	}
	else {
		return true
	}
}

PreventiceWebPager(phase,chgStr,btnStr) {
	pg0 := gl.Page.getElementById(chgStr).innerText
	
	if (tot0 := stRegX(pg0,"i)items.*? of ",1,1,"\d+",0)) {
		gl.inv_tot := tot0
	}

	if (phase="Enrollment") {
		pgNum := gl.enroll_ct
		gl.Page.getElementById(btnStr).click() 											; click when id=btnStr
	}
	if (phase="Inventory") {
		pgNum := gl.inv_ct
		progCt := gl.inv_tot
		if (gl.Page.getElementsByClassName(btnStr)[0].getAttribute("onClick") ~= "return") {
			return
		}
		gl.Page.getElementsByClassName(btnStr)[0].click() 								; click when class=btnstr
	}
	
	t0 := A_TickCount
	While (A_TickCount-t0 < gl.settings.webwait)
	{
		pb.set(100*pgNum/progCt)
		
		pg := gl.Page.getElementById(chgStr).innerText
		if (pg != pg0) {
			t1:=A_TickCount-t0
			eventlog(phase " " pgNum " pager (" round(t1/1000,2) " s)"
					, (t1>5000) ? 1 : 0)
			return
		}
		sleep 50
	}
	eventlog(phase " " pgNum " timed out! (" round((A_TickCount-t0)/1000,2) " s)")
	return
}

parsePreventiceInventory(tbl) {
/*	Parse Preventice website for device inventory
	Add unique ser nums to /root/inventory/dev[@ser]
	These will be removed when registered
*/
	gl.clip := tbl.innerText
	if (gl.clip=gl.clip0) {																; no change since last clip
		Return false
	}

	lbl := ["button","model","ser"]
	
	trows := tbl.getElementsbyTagName("tr")
	loop trows.count																	; loop through rows
	{
		r_idx := A_index-1
		trow := trows[r_idx]
		tcols := trow.getElementsbyTagName("td")
		res := Map()
		loop lbl.length																	; loop through cols
		{
			c_idx := A_Index-1
			res.%lbl[A_index]% := trim(tcols[c_idx].innerText)
		}
		gl.inv_ct++
		
		if IsObject(wq.selectSingleNode("/root/pending/enroll"
					. "[dev='" res.model " - " res.ser "']")) {							; exists in Pending
			eventlog(res.model " - " res.ser " - already in use.",0)
			continue
		}
		
		gl.prevtxt .= "dev|" res.model "|" res.ser "`n"
	}
	gl.clip0 := gl.clip																	; set the check for repeat copy

	return true
}



;#endregion

;#region == TEXT Elements ==============================================================
eventlog(event,verbosity:=1) {
	/*	verbose 1 or 0 from ini
		verbosity default 1
		verbosity set 0 if only during verbose
	*/
		score := verbosity + gl.settings.verbose
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
			k := &key
			v := &val
			%k% := %v%
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


;#endregion

#Include xml2.ahk
#Include strx2.ahk
#Include %A_ScriptDir%\Rufaydium.ahk
