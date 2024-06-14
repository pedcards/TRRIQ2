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
	gl.adminMode := false

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
	gl.isMain := (wksLoc~="Main Campus") ? true : false

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

	; pb.sub("HL7 map")
	; hl7 := getHL7()																						; HL7 definitions
	; hl7DirMap := {}

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
	; saveCygnusLogs("all")

	pb.title("Cleaning old .bak files")
	pb.sub("")
	cleanBakFiles()
	pb.Hide

		
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
	global gl, dims, sites, wksLoc
		, phase, WQtab
	
	dims.hwnd := Map()
	phase := Gui()
	phase.Opt("+AlwaysOnTop")

	/*	Phase info box
	 */
	phaseNumbers := phase.AddText("x" dims.phase.lvW+40 " y15 w200 vPhaseNumbers", "`n`n")
	phase.AddGroupBox("x" dims.phase.lvW+20 " y0 w220 h65")

	/*	Action buttons
	 */
	phase.SetFont("Bold","Verdana")
	btnRefresh := phase.AddButton("Y+10 wp h40","Refresh lists") ; gPhaseRefresh
		; btnRefresh.OnEvent("Click",PhaseRefresh())
	btnPrevGrab := phase.AddButton("Y+10 wp h40 Disabled","Check Preventice inventory") ; gPrevGrab
		btnPrevGrab.OnEvent("Click",prevgrab)
	phase.AddText("wp h50")
	phase.AddText("y+10 wp h24 Center","Register/Prepare a `nHOLTER or EVENT MONITOR")
	btnOrders := phase.AddButton("y+10 wp h40 vRegister DISABLED","No active orders") ; gPhaseOrder
		btnOrders.OnEvent("Click",phaseOrder)
		dims.hwnd["btnOrders"] := btnOrders.Hwnd
	phase.AddText("wp h30")
	phase.AddText("y+10 wp Center","Transmit")
	btnBGM := phase.AddText("y+1 wp Center h100","BG MINI")
		btnBGM.GetPos(&bgmX,&bgmY,&bgmW,&bgmH)
		; btnBGM.OnEvent("Click",holterUpload)
	phase.SetFont("norm")

	/* 	BG MINI button
	 */
	btnW := 79
	btnH := 61
	btnBGM := phase.AddPicture("y" bgmY+20 " x" bgmX+70 " w" bgmW/3 " h" bgmH/2
		,".\files\BGMini.png")
	
	/*	MAIN TABVIEW
	 */
	phase.SetFont(,"Calibri")
	siteTabs := makeSiteTab()
	WQtab := phase.AddTab3("x10 y10 w" dims.phase.lvW " h" dims.phase.lvH 
		; . " +HwndWQtab -Wrap"
		, siteTabs)
	WQtab.GetPos(&wqX,&wqY,&wqW,&wqH)
		dims.wqTab := Map()
		dims.wqTab.X := wqX
		dims.wqTab.Y := wqY
		dims.wqTab.W := wqW
		dims.wqTab.H := wqH
		dims.hwnd["WQtab"] := WQtab.Hwnd
	
	/*	BUILD LISTVIEWS
	 */
	lvDim := "w" wqW-25 " h" wqH-35

	if (gl.isMain) {
		btnPrevGrab.Enabled := true

		WQtab.UseTab("INBOX") ; ======================================================== INBOX
		HLV_in := phase.AddListView("-Multi Grid BackgroundSilver " lvDim
			, ["filename","Name","MRN","DOB","Location","Study Date","wqid","Type","Need FTP"]
		)
		dims.hwnd["HLV_in"] := HLV_in.Hwnd
		; HLV_in.OnEvent("DoubleClick",readWQlv())
		HLV_in.ModifyCol(1,"0")															; filename and path, "0" = hidden
		HLV_in.ModifyCol(2,"160 Center")												; name
		HLV_in.ModifyCol(3,"60 Center")													; mrn
		HLV_in.ModifyCol(4,"80 Center")													; dob
		HLV_in.ModifyCol(5,"80 Center")													; site
		HLV_in.ModifyCol(6,"80 Center")													; date
		HLV_in.ModifyCol(7,"2")															; wqid
		HLV_in.ModifyCol(8,"40 Center")													; ftype
		HLV_in.ModifyCol(9,"70 Center")													; ftp
		; CLV_in := new LV_Colors(HLV_in,true,false)
		; CLV_in.Critical := 100
	}

	WQtab.UseTab("ORDERS") ; =========================================================== ORDERS
	HLV_orders := phase.AddListView("-Multi Grid BackgroundSilver " lvDim	; option "ColorRed"
		, ["filename","Order Date","Name","MRN","Ordering Provider","Monitor"]
	)
	dims.hwnd["HLV_orders"] := HLV_orders.Hwnd
	; HLV_orders.OnEvent("DoubleClick",readWQorder())
	HLV_orders.ModifyCol(1,"0")															; filename and path (hidden)
	HLV_orders.ModifyCol(2,"80")														; date
	HLV_orders.ModifyCol(3,"140")														; Name
	HLV_orders.ModifyCol(4,"60")														; MRN
	HLV_orders.ModifyCol(5,"100")														; Prov
	HLV_orders.ModifyCol(6,"70")														; Type
	
	WQtab.UseTab("Unread") ; =========================================================== UNREAD
	HLV_unread := phase.AddListView("-Multi Grid BackgroundSilver " lvDim
		, ["Name","MRN","Study Date","Processed","Monitor","Ordering","Assigned EP"]
	)
	HLV_unread.ModifyCol(1,"140")														; Name
	HLV_unread.ModifyCol(2,"60")														; MRN
	HLV_unread.ModifyCol(3,"80")														; Date
	HLV_unread.ModifyCol(4,"80")														; Processed
	HLV_unread.ModifyCol(5,"70")														; Mon Type
	HLV_unread.ModifyCol(6,"80")														; Ordering
	HLV_unread.ModifyCol(7,"80")														; Assigned EP

	WQtab.UseTab("ALL") ; ============================================================== ALL
	HLV_all := phase.AddListView("-Multi Grid BackgroundSilver " lvDim
		, ["ID","Enrolled","FedEx","Uploaded","Notes","MRN","Enrolled Name","Device","Provider","Site"]
	)
	; HLV_all.OnEvent("DoubleClick",WQtask())
	HLV_all.ModifyCol(1,"0")															; wqid (hidden)
	HLV_all.ModifyCol(2,"60")															; date
	HLV_all.ModifyCol(3,"40 Center")													; FedEx
	HLV_all.ModifyCol(4,"60")															; uploaded
	HLV_all.ModifyCol(5,"40 Center")													; Notes
	HLV_all.ModifyCol(6,"60")															; MRN
	HLV_all.ModifyCol(7,"140")															; Name
	HLV_all.ModifyCol(8,"130")															; Ser Num
	HLV_all.ModifyCol(9,"100")															; Prov
	HLV_all.ModifyCol(10,"80")															; Site
	; CLV_all := new LV_Colors(HLV_all,true,false)
	; CLV_all.Critical := 100

	; ================================================================================== LV for each Site
	HLV1 := HLV2 := HLV3 := HLV4 := HLV5 := HLV6 := HLV7 := HLV8 := HLV9 := ""			; Must declare first, V2 cannot create dynamic variable names
	loop parse sites.tracked, "|"
	{
		i := A_Index
		site := A_LoopField
		WQtab.UseTab(site)
		HLV%i% := phase.AddListView("-Multi Grid BackgroundSilver " lvDim
			, ["ID","Enrolled","FedEx","Uploaded","Notes","MRN","Enrolled Name","Device","Provider"]
		)
		; HLV%i%.OnEvent("DoubleClick",WQtask())
		HLV%i%.ModifyCol(1,"0")															; wqid (hidden)
		HLV%i%.ModifyCol(2,"60")														; date
		HLV%i%.ModifyCol(3,"40 Center")													; FedEx
		HLV%i%.ModifyCol(4,"60")														; uploaded
		HLV%i%.ModifyCol(5,"40 Center")													; Notes
		HLV%i%.ModifyCol(6,"60")														; MRN
		HLV%i%.ModifyCol(7,"140")														; Name
		HLV%i%.ModifyCol(8,"130")														; Ser Num
		HLV%i%.ModifyCol(9,"100")														; Prov
		; CLV_%i% := new LV_Colors(HLV%i%,true,false)
		; CLV_%i%.Critical := 100
	}

	/*	POPULATE LISTVIEWS
	 */
	WQlist()
	
	/*	MENUS
	 */
	menuSys := Menu()
		menuSys.Add("Change clinic location", changeLoc)
		menuSys.Add("Generate late returns report", menuAbout) ;,lateReport())
		menuSys.Add("Generate registration locations report", menuAbout) ;,regReport())
		menuSys.Add("Update call schedules", menuAbout) ;, updateCall())
	menuHelp := Menu()
		menuHelp.Add("About TRRIQ", menuAbout)
		menuHelp.Add("Instructions...", menuInstructions)
	menuAdmin := Menu()
		menuAdmin.Add("Toggle admin mode", toggleAdmin)
		menuAdmin.Add("Clean tempfiles", menuAbout) ;, CleanTempFiles())
		menuAdmin.Add("Send notification email", menuAbout) ;, sendEmail())
		menuAdmin.Add("Find pending leftovers", menuAbout) ;, cleanPending())
		menuAdmin.Add("Fix WQ device durations", menuAbout) ;, fixDuration())							; position for test menu
		menuAdmin.Add("Recover DONE record", menuAbout) ;, recoverDone())
		menuAdmin.Add("Check running users/versions", menuAbout) ;, runningUsers())
		menuAdmin.Add("Create test order", menuAbout) ;, makeEpicORM())
		
	phaseMenu := MenuBar()
		phaseMenu.Add("System",menuSys)
		if (gl.user~="i)tchun1|docte") {
			phaseMenu.Add("Admin",menuAdmin)
		}
		phaseMenu.Add("Help",menuHelp)
	
	phase.MenuBar := phaseMenu

	if (gl.adminMode) {
		phase.BackColor := "Fuchsia"
		phase.Title := "TRRIQ Dashboard - ADMIN MODE"
	} else {
		phase.Title := "TRRIQ Dashboard"
	}
	phase.Show()
	phase.OnEvent("Close",phaseClose)
	/*
		
		SetTimer, idleTimer, 500
		return
	*/

	RETURN

	/*	Internal PhaseGUI methods ======================================================
	 */
	phaseClose(*) {
		ask := MsgBox("Really quit TRRIQ?","Exit",262161)
		If (ask="OK")
		{
			; checkPreventiceOrdersOut()
			; cleanDone()
			eventlog("<<<<< Session end.")
			ExitApp
		}
	}	

	makeSiteTab() {
		alltabs := 
			["ORDERS"
			, "INBOX"
			, "Unread"
			, "ALL"
			]
		if !(gl.isMain) {																	; If not MAIN, remove INBOX
			alltabs.RemoveAt(2)
		}
		for val in StrSplit(sites.tracked,"|")
		{
			alltabs.Push(val)
		}
		return alltabs
	}

	prevGrab(*) {
		Run("prevgrab.exe")
		return
	}

	PhaseOrder(*) {
		WQtab.Choose("ORDERS")
		return
	}

	menuAbout(*) {
		phase.Hide
		MsgBox(A_ScriptName " version " substr(gl.runningVer,1,12) "`nTerrence Chun, MD","About...",64)
		phase.Show
		return
	}

	menuInstructions(*) {
		phase.Hide
		MsgBox "How to..."
		phase.Show
		return
	}

	changeLoc(*) {
		ask := MsgBox("Current location: " wksLoc "`n`nReally change the clinic location for this PC?`n`nWill restart TRRIQ"
			, "Change clinic", 262193)
		If (ask="OK")
		{
			wks.clearWorkstation()
			Reload
		}
		return
	}

	toggleAdmin(*) {
		gl.adminMode := !(gl.adminMode)
		PhaseGUI()
		return
	}


}

WQlist() {
/*	Build up listviews from file sources
	Read and update worklist => wq
 */
	global wq, dims, phase, WQtab
	
	wqfiles := []
	fldval := {}

	lvDim := "w" dims.wqTab.W-25 " h" dims.wqTab.H-35
	
	checkversion()																		; make sure we are running latest version

	pb.title("Scanning worklist...")
	
	fileCheck()
	FileOpen(".lock", "W")																; Create lock file.
	
	wq := XML(path.data "worklist.xml")													; refresh WQ
	
	readPrevTxt()																		; read prev.txt from website
	
	WQclearSites0()	 																	; move studies from sites0 to DONE
	
	/*	Add all incoming Epic ORDERS to WQlv_orders
	*/
	lv := GuiCtrlFromHwnd(dims.hwnd["HLV_orders"])
	lv.Delete()
	
	pb.title("Scanning Epic orders")
	pb.set()
	WQscanEpicOrders(lv)
	
	; WriteSave(wq)
	; FileDelete(".lock")
	
	; checkPreventiceOrdersOut()															; check registrations that failed upload to Preventice
	
	/*	Generate Inbox WQlv_in tab for Main Campus user 
	/
	if (gl.isMain) {
		Gui, ListView, WQlv_in
		LV_Delete()																		; clear the INBOX entries
		
		WQpreventiceResults(wqfiles)													; Process incoming Preventice results
		WQscanHolterPDFs(wqfiles)														; Scan Holter PDFs folder for additional files
		WQfindMissingWebgrab()															; find <pending> missing <webgrab>
	}
	
	/*	Generate lv for ALL, site tabs, and pending reads
	/
	WQpendingTabs()

	WQpendingReads()

	GuiControl, Text, PhaseNumbers
		,	% "Patients registered in Preventice (" wq.selectNodes("/root/pending/enroll").length ")`n"
		.	(tmp := parsedate(wq.selectSingleNode("/root/pending").getAttribute("update")))
		.	"Preventice update: " tmp.MMDD " @ " tmp.hrmin "`n"
		.	(tmp := parsedate(wq.selectSingleNode("/root/inventory").getAttribute("update")))
		.	"Inventory update: " tmp.MMDD " @ " tmp.hrmin
	
	progress, off
	return
*/
}


;#endregion

;#region == FILE handling elements =====================================================
checkVersion() {
/*	Checks running version
	In event user has not restarted TRRIQ since last update
 */
	chk := FileGetTime(A_ScriptName)
	if (chk != gl.runningVer) {
		ask := MsgBox("There is an updated version of the script. `nRestart to launch new version?"
			"New version!", 262193)
		If (ask="OK")
			run A_ScriptName
		ExitApp
	}
	return
}

parseORM() {
/*	parse fldval values to values
	including aliases for both WQlist and readWQorder
*/
	global fldval, sites, indCodes
	
	monType:=(tmp:=fldval["OBR_TestName"])~="i)14 DAY" ? "BGM"							; for extended recording
		: tmp~="i)15 DAY" ? "BGM"
		: tmp~="i)24 HOUR" ? "HOL"														; for short report (includes full disclosure)
		: tmp~="i)48 HOUR" ? "HOL"
		: tmp~="i)RECORDER|EVENT" ? "BGH"
		: tmp~="i)CUTOVER" ? "CUTOVER"
		: ""
	
	switch fldval["PV1_PtClass"]
	{
		case "O":
			encType := "Outpatient"
			location := sites.Long[fldval["PV1_Location"]]
		case "I":
			encType := "Inpatient"
			location := "MAIN"
		case "OBS":
			encType := "Inpatient"
			location := "MAIN"
		case "DS":
			encType := "Outpatient"
			location := "MAIN"
		case "E":
			encType := "Inpatient"
			location := "Emergency"
		default:
			encType := "Outpatient"
			location := fldval["PV1_Location"]
	}
	prov := strQ(fldval["ORC_ProvCode"]
			, fldval["ORC_ProvCode"] "^" fldval["ORC_ProvNameL"] "^" fldval["ORC_ProvNameF"]
			, fldval["OBR_ProviderCode"] "^" fldval["OBR_ProviderNameL"] "^" fldval["OBR_ProviderNameF"])
	provname := strQ(fldval["ORC_ProvCode"]
			, fldval["ORC_ProvNameL"] strQ(fldval["ORC_ProvNameF"], ", ###")
			, fldval["OBR_ProviderNameL"] strQ(fldval["OBR_ProviderNameF"], ", ###"))
	provHL7 := fldval["ORC"][12]
	;~ location := (encType="Outpatient") ? sitesLong[fldval.PV1_Location]
		;~ : encType
	
	indText := indCode := indication := ""
	if !(indication:=strQ(fldval["OBR_ReasonCode"],"###") strQ(fldval["OBR_ReasonText"],"^###")) {
		indText := objhasvalue(fldval,"^Reason for exam","RX")
		indText := (indText="hl7") ? "" : indText										; no "Reason for exam" returns "hl7", breaks fldval[indtext]
		indText := RegExReplace(fldval[indText],"Reason for exam->")
		
		indCode := objhasvalue(indCodes,indText,"RX")
		indCode := strX(indCodes[indCode],"",1,0,":",1,1)
		
		indication := strQ(indCode,"###") strQ(indText,"^###")
	}
	
	return {date:parseDate(fldval["OBR_StartDateTime"]).YMD
		, encDate:parseDate(fldval["PV1_DateTime"]).YMD
		, namePID5:fldval["PID"][5]
		, nameL:fldval["PID_NameL"]
		, nameF:fldval["PID_NameF"]
		, name:fldval["PID_NameL"] strQ(fldval["PID_NameF"],", ###")
		, mrn:fldval["PID_PatMRN"]
		, sex:(fldval["PID_Sex"]~="F") ? "Female" : (fldval["PID_Sex"]~="M") ? "Male" : (fldval["PID_Sex"]~="U") ? "Unknown" : ""
		, DOB:parseDate(fldval["PID_DOB"]).MDY
		, monitor:monType
		, mon:monType
		, provider:prov
		, prov:prov
		, provname:provname
		, provORC12:provHL7
		, type:encType
		, loc:location
		, Account:fldval["ORC_ReqNum"]
		, accountnum:fldval["PID_AcctNum"]
		, encnum:fldval["PV1_VisitNum"]
		, order:fldval["ORC_ReqNum"]
		, accession:fldval["ORC_FillerNum"]
		, UID:tobase(fldval["ORC_ReqNum"] RegExReplace(fldval["ORC_FillerNum"],"[^0-9]"),36)
		, ind:indication
		, indication:indication
		, indicationCode:strQ(fldval["OBR_ReasonCode"],"###") strQ(indCode,"###")
		, orderCtrl:fldval["ORC_OrderCtrl"]
		, ctrlID:fldval["MSH_CtrlID"]}
}
	
;#endregion

;#region == TEXT functions =============================================================
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
	try return (var1="") ? null : RegExReplace(txt,"###",var1)
}

filterProv(x) {
/*	Filters out all irregularities and common typos in Provider name from manual entry
	Returns as {name:"Albers, Erin", site:"CRB"}
	Provider-Site may be in error
*/
	global sites
	
	allsites := sites.tracked "|" sites.ignored
	RegExMatch(x,"i)-(" allsites ")\s*,",&site)
	x := trim(x)																		; trim leading and trailing spaces
	x := RegExReplace(x,"i)\s{2,}"," ")													; replace extra spaces
	x := RegExReplace(x,"i)\s*-\s*(" allsites ")$")										; remove trailing "LOUAY TONI(-tri)"
	x := RegExReplace(x,"i)( [a-z](\.)? )"," ")											; remove middle initial "STEPHEN P SESLAR" to "Stephen Seslar"
	x := RegExReplace(x,"i)^Dr(\.)?(\s)?")												; remove preceding "(Dr. )Veronica..."
	x := RegExReplace(x,"i)^[a-z](\.)?\s")												; remove preceding "(P. )Ruggerie, Dennis"
	x := RegExReplace(x,"i)\s[a-z](\.)?$")												; remove trailing "Ruggerie, Dennis( P.)"
	x := RegExReplace(x,"i)\s*-\s*(" allsites ")\s*,",",")								; remove "SCHMER(-YAKIMA), VERONICA"
	x := RegExReplace(x,"i) (MD|DO)$")													; remove trailing "( MD)"
	x := RegExReplace(x,"i) (MD|DO),",",")												; replace "Ruggerie MD, Dennis" with "Ruggerie, Dennis"
	x := RegExReplace(x," NPI: \d{6,}$")												; remove trailing " NPI: xxxxxxxxxx"
	x := StrTitle(x)																	; convert "RUGGERIE, DENNIS" to "Ruggerie, Dennis"
	if !InStr(x,", ") {
		x := strX(x," ",1,1,"",1,0) ", " strX(x,"",1,1," ",1,1)							; convert "DENNIS RUGGERIE" to "RUGGERIE, DENNIS"
	}
	x := RegExReplace(x,"^, ")															; remove preceding "(, )Albers" in event this happens
	if (site="") {
		site := Map()
		site[1] := "MAIN"
	}
	if (site[1]="TRI") {																; sometimes site improperly registered as "tri"
		site[1] := "TRI-CITIES"
	}
	return {name:x, site:site[1]}
}

	
;#endregion

;#region == WQ/Worklist FUNCTIONS ======================================================
readWQ(idx) {
	global wq
	
	res := {}
	k := wq.selectSingleNode("//enroll[@id='" idx "']")
	Loop (ch:=k.selectNodes("*")).Length
	{
		i := ch.item(A_Index-1)
		node := i.nodeName
		val := i.text
		res.%node% := val
	}
	res.node := k.parentNode.nodeName 
	
	return res
}

wqSetVal(id,node,val) {
	global wq
	
	newID := "/root/pending/enroll[@id='" id "']"
	k := wq.selectSingleNode(newID "/" node)
	if (k.text) and (val="") {															; don't overwrite an existing value with null
		return
	}
	val := RegExReplace(val,"\'","^")													; make sure no val ever contains [']
	
	if IsObject(k) {
		k.text := val
	} else {
		wq.addElement(newID,node,val)
	}
	
	return
}

checkweb(id) {
	global wq

	en := "//enroll[@id='" id "']"
	if (wq.selectSingleNode(en "/webgrab").text) {											; webgrab already exists
		Return
	} else {
		wq.addElement(en,"webgrab",A_Now)
		eventlog("Added webgrab for id " id)
		Return
	}
}

makeUID() {
	global wq
	
	Loop
	{
		num1 := Random(10000, 99999)
		num2 := Random(10000, 99999)
		num3 := Random(10000, 99999)
		num := num1 . num2 . num3
		id := toBase(num,36)
		if IsObject(wq.selectSingleNode("//enroll[id='" id "']")) {
			eventlog("makeUID: " id " already in use.")
			continue
		} 
		else {
			break
		}
	}
	return id
}

WQclearSites0() {
/*	Clear enroll nodes from sites0 locations
*/
	global sites, wq

	loop parse sites.ignored, "|"
	{
		site := A_LoopField
		Loop (ens:=wq.selectNodes("/root/pending/enroll[site='" site "']")).length
		{
			k := ens.item(A_Index-1)
			clone := k.cloneNode(true)
			wq.selectSingleNode("/root/done").appendChild(clone)						; copy k.clone to DONE
			k.parentNode.removeChild(k)													; remove k node
			eventlog("Moved " site " record " k.selectSingleNode("mrn").text " " k.selectSingleNode("name").text)
		}
	}
	Return
}

WQscanEpicOrders(lv) {
/*	Scan all incoming Epic orders
	3-pass method
*/
	global wq

	if !IsObject(wq.selectSingleNode("/root/orders")) {
		wq.addElement("/root","orders")
	}
	
	WQEpicOrdersNew()																	; Process new files

	WQEpicOrdersPrevious()																; Scan previous *Z.hl7 files

	; WQepicOrdersCleanup()																; Remove extraneous orders

	lv.ModifyCol(2,"SortDesc")															; Sort orders LV by date

	Return
}

WQepicOrdersNew() {
/*	First pass: process new files
	Find noval (not renamed) hl7 files in path.EpicHL7in
	Find matching <enroll> node
		Skip sites0
		Skip Name or MRN string varies by more than 15%
		Skip datediff > 5d
	Adjust name, order, accession, account, encounter num for <enroll> node
	Handle corresponding <orders> node
*/
	global wq, path, sites, fldVal
	pb.sub("New orders...")

	Loop files path.EpicHL7in "*"
	{
		pb.set(A_Index*4)
		e0 := {}
		fileIn := A_LoopFileName
		if RegExMatch(fileIn,"_@([a-zA-Z0-9]{4,}).hl7") {								; skip old files
			continue
		}
		ord_in := hl7(A_LoopFileFullPath)
		fldval := ord_in.fldval
		e0:=parseORM()
		if InStr(sites.ignored, e0.loc) {												; skip non-tracked orders
			FileMove(A_LoopFileFullPath, ".\tempfiles\" e0.mrn "_" e0.nameL "_" A_LoopFileName, 1)
			eventlog("Non-tracked order " fileIn " moved to tempfiles. " e0.loc " " e0.mrn " " e0.nameL)
			continue
		}
		eventlog("New order " fileIn ". " e0.name " " e0.mrn )
		
		loop (ens:=wq.selectNodes("/root/pending/enroll")).Length						; find enroll nodes with result but no order
		{
			k := ens.item(A_Index-1)
			if IsObject(k.selectSingleNode("accession")) {								; skip nodes that already have accession
				continue
			}
			pb.sub(k.selectSingleNode("name").text)
			e0.match_NM := fuzzysearch(e0.name,format("{:U}",k.selectSingleNode("name").text))
			e0.match_MRN := fuzzysearch(e0.mrn,k.selectSingleNode("mrn").text)
			if (e0.match_NM > 0.15) || (e0.match_MRN > 0.15) {							; Name or MRN vary by more than 15%
				continue
			}
			dt0 := dateDiff(e0.date,k.selectSingleNode("date").text,"Days")
			if abs(dt0) > 5 {															; Date differs by more than 5d
				Continue
			}

			id := k.getAttribute("id")
			e0.match_UID := true
			
			if (e0.name != k.selectSingleNode("name").text) {
				wqSetVal(id,"name",e0.name)
				eventlog("enroll name " k.selectSingleNode("name").text " changed to " e0.name)
			}
			wqSetVal(id,"order",e0.order)
			wqSetVal(id,"accession",e0.accession)
			wqSetVal(id,"acctnum",e0.accountnum)
			wqSetVal(id,"encnum",e0.encnum)	
			k.setAttribute("id",e0.UID)
			eventlog("Found pending/enroll=" id " that matches new Epic order " e0.order ". " e0.match_NM)
			eventlog("enroll id " id " changed to " e0.UID)
			break
		}
		try if (e0.match_UID) {
			FileMove(A_LoopFileFullPath, ".\tempfiles\*", 1)
			eventlog("Moved: " A_LoopFileFullPath)
			continue
		}
		
		e0.orderNode := "/root/orders/enroll[order='" e0.order "']"
		if IsObject(k:=wq.selectSingleNode(e0.orderNode)) {								; ordernum node exists
			e0.nodeCtrlID := k.selectSingleNode("ctrlID").text
			if (e0.CtrlID < e0.nodeCtrlID) {											; order CtrlID is older than existing, somehow
				FileDelete(path.EpicHL7in fileIn)
				eventlog("Order msg " fileIn " is outdated. " e0.name)
				continue
			}
			if (e0.orderCtrl="CA") {													; CAncel an order
				FileDelete(path.EpicHL7in fileIn)										; delete this order message
				FileDelete(path.EpicHL7in "*_@" e0.UID ".hl7")							; and the previously processed hl7 file
				wq.removeNode(e0.orderNode)												; and the accompanying node
				eventlog("Cancelled order " e0.order ". " e0.name)
				continue
			}
			FileDelete(path.EpicHL7in "*_@" e0.UID ".hl7")								; delete previously processed hl7 file
			wq.removeNode(e0.orderNode)													; and the accompanying node
			eventlog("Cleared order " e0.order " node. " e0.name)
		}
		if (e0.orderCtrl="XO") {														; change an order
			e0.orderNode := "/root/orders/enroll[accession='" e0.accession "']"
			k := wq.selectSingleNode(e0.orderNode)
			e0.nodeUID := k.getAttribute("id")
			FileDelete(path.EpicHL7in "*_@" e0.nodeUID ".hl7")
			wq.removeNode(e0.orderNode)
			eventlog("Removed node id " e0.nodeUID " for replacement. " e0.name)
		}
		
		newID := "/root/orders/enroll[@id='" e0.UID "']"								; otherwise create a new node
			wq.addElement("/root/orders","enroll",{id:e0.UID})
			wq.addElement(newID,"order",e0.order)
			wq.addElement(newID,"accession",e0.accession)
			wq.addElement(newID,"ctrlID",e0.CtrlID)
			wq.addElement(newID,"date",e0.date)
			wq.addElement(newID,"name",e0.name)
			wq.addElement(newID,"mrn",e0.mrn)
			wq.addElement(newID,"sex",e0.sex)
			wq.addElement(newID,"dob",e0.dob)
			wq.addElement(newID,"mon",e0.mon)
			wq.addElement(newID,"prov",e0.prov)
			wq.addElement(newID,"provname",e0.provname)
			wq.addElement(newID,"site",e0.loc)
			wq.addElement(newID,"acctnum",e0.accountnum)
			wq.addElement(newID,"encnum",e0.encnum)
			wq.addElement(newID,"ind",e0.ind)
		eventlog("Added order ID " e0.UID ". " e0.name)
		
		fileOut := (e0.mon="CUTOVER" ? "done\" : "")
			. e0.MRN "_" 
			. fldval["PID_NameL"] "^" fldval["PID_NameF"] "_"
			. e0.date "_@"
			. e0.uid 																	; new ORM filename ends with _[UID]Z.hl7
			. ".hl7"
		
		FileMove(A_LoopFileFullPath, path.EpicHL7in . fileOut)							; and rename ORM file
	}

	Return
}

WQepicOrdersPrevious() {
/*	Second pass: scan previously added *Z.hl7 files
	Another chance to clear sites0 and remnant files
	Add line to Inbox LV
*/
	global path, wq, monTypes, dims, phase

	loop Files path.EpicHL7in "*_@*.hl7"
	{
		e0 := {}
		fileIn := A_LoopFileName
		if RegExMatch(fileIn,"_@([a-zA-Z0-9]{4,}).hl7",&i) {							; file appears to have been parsed
			e0 := readWQ(i[1])
		} else {
			continue
		}
		
		if InStr(sites.ignored,e0.site) {												; sites0 location
			FileMove(A_LoopFileFullPath, ".\tempfiles", 1)
			wq.removeNode("/root/orders/enroll[@id='" i[1] "']")
			eventlog("Non-tracked order " fileIn " moved to tempfiles.")
			continue
		}
		if (e0.node ~= "pending|done") {												; remnant orders file
			FileMove(A_LoopFileFullPath, ".\tempfiles", 1)
			eventlog("Leftover HL7 file " fileIn " moved to tempfiles.")
			continue
		}
		
		lv := GuiCtrlFromHwnd(dims.hwnd["HLV_orders"])
		lv.Add(""
			, path.EpicHL7in . fileIn													; filename and path to HolterDir
			, e0.date																	; date
			, e0.name																	; name
			, e0.mrn																	; mrn
			, e0.provname																; prov
			, getMonType(e0.mon)["abbrev"] " "												; monitor type
				; . (e0.mon="BGH"															; relabel BGH=>CEM
				; ? "CEM"
				; : e0.mon)
			, "")																		; fulldisc present, make blank
		btn := GuiCtrlFromHwnd(dims.hwnd["btnOrders"])
			btn.Enabled := true
			btn.Text := "Go to ORDERS tab"
	}
	Return
}

WQepicOrdersCleanup() {
/*	Third pass: remove extraneous orders
*/
/*
	global wq

	loop, % (ens:=wq.selectNodes("/root/orders/enroll")).Length
	{
		e0 := {}
		k := ens.item(A_Index-1)
		e0.uid := k.getAttribute("id")
		e0.order := k.selectSingleNode("order").text
		e0.accession := k.selectSingleNode("accession").text
		e0.name := k.selectSingleNode("name").text
		
		if IsObject(wq.selectSingleNode("/root/pending/enroll[order='" e0.order "'][accession='" e0.accession "']")) {
			eventlog("Order node " e0.uid " " e0.name " already found in pending.")
			removenode("/root/orders/enroll[@id='" e0.uid "']")
		}
		if IsObject(wq.selectSingleNode("/root/done/enroll[order='" e0.order "'][accession='" e0.accession "']")) {
			eventlog("Order node " e0.uid " " e0.name " already found in done.")
			removenode("/root/orders/enroll[@id='" e0.uid "']")
		}
	}
	Return
*/
}

getMonType(val) {
	res := 0
	for key,arr in monTypes
	{
		if ObjHasValue(arr,val) {
			res := A_Index
			break
		}
	}
	try return monTypes[res]
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

readPrevTxt() {
/*	Read data files from Preventice:
		* Patient Status Report_v2.xml sent by email every M-F 6 AM
		* prev.txt grabbed from prevgrab.exe
			- Enrollments (inactive, as taken from PSR_v2)
			- Inventory
*/
	global wq
	
	pb.title("Updating Patient Status Report data")

	psr := XML(path.data "Patient Status Report_v2.xml")
		psrdate := parseDate(psr.selectSingleNode("Report").getAttribute("ReportTitle"))	; report date is in Central Time
		psrDT := psrdate.YMDHMS
	psrlastDT := wq.selectSingleNode("/root/pending").getAttribute("update")
	if (psrDT>psrlastDT) {																; check if psrDT more recent
		pb.sub("Reading registration updates...")
		dets := psr.selectNodes("//Details_Collection/Details")
		numdets := dets.length()
		loop numdets
		{
			pb.set(A_Index)
			k := dets.item(numdets-A_Index)												; read nodes from oldest to newest
			parsePrevEnroll(k)
		}
		wq.selectSingleNode("/root/pending").setAttribute("update",psrDT)				; set pending[@update] attr
		eventlog("Patient Status Report " psrDT " updated.")

		; lateReportNotify()
	}

	filenm := path.data "prev.txt"
	filedt := FileGetTime(filenm)
	lastInvDT := wq.selectSingleNode("/root/inventory").getAttribute("update")
	if (filedt=lastInvDT) {
		Return
	}
	pb.sub("Reading inventory updates...")
	txt := FileRead(filenm)
	txt := StrReplace(txt, "`n", "`n",, &n)		 										; count number of lines
	devct := false
	
	loop parse txt, "`r`n", "`r"
	{
		pb.set(100*A_Index/n)
		
		k := A_LoopField
		if (k~="^dev\|") {
			if !(devct) {
				inv := wq.selectSingleNode("/root/inventory")							; create fresh inventory node
				inv.parentNode.removeChild(inv)
				wq.addElement("/root","inventory")
				devct := true
			}
			parsePrevDev(k)
		}
	}
	
	loop (devs := wq.selectNodes("/root/inventory/dev")).length							; Find dev that already exist in Pending
	{
		k := devs.item(A_Index-1)
		dev := k.getAttribute("model")
		ser := k.getAttribute("ser")
		if IsObject(wq.selectSingleNode("/root/pending/enroll[dev='" dev " - " ser "']")) {	; exists in Pending
			k.parentNode.removeChild(k)
			eventlog("Removed inventory ser " ser)
		}
	}
	wq.selectSingleNode("/root/inventory").setAttribute("update",filedt)				; set pending[@update] attr
	eventlog("Preventice Inventory " fileDT " updated.")
	
return	
}

parsePrevEnroll(det) {
/*	Parse line from Patient Status Report_v2
	"enroll"|date|name|mrn|dev - s/n|prov|site
	Match to existing/likely enroll nodes
	Update enroll node with new info if missing
*/
	global wq, sites

	res := {  date:parseDate(det.getAttribute("Date_Enrolled")).YMD
			, name:RegExReplace(format("{:U}"
					,det.getAttribute("PatientLastName") ", " det.getAttribute("PatientFirstName"))
					,"\'","^")
			, mrn:det.getAttribute("MRN1")
			, dev:det.getAttribute("Device_Type") " - " det.getAttribute("Device_Serial")
			, prov:filterProv(det.getAttribute("Ordering_Physician")).name
			, site:filterProv(det.getAttribute("Ordering_Physician")).site
			, id:det.getAttribute("CSN_SecondaryID1") 
			, duration:det.getAttribute("Study_Duration") }

	if InStr(res.name,"`"") {
		res.name := trim(RegExReplace(res.name,"\`".*?\`""))							; delete "quoted" nicknames
	}
	if (res.dev~=" - $") {																; e.g. "Body Guardian Mini -"
		res.dev .= res.name																; append string so will not match in enrollcheck
	}
	
	/*	Ignore sites.ignored enrollments entirely
	*/
		if (res.site~=sites.ignored) {
			Return
		}

	/*	Check whether any params match this device
	*/
		if (id:=enrollcheck("[@id='" res.id "']")) {									; id returned in Preventice ORU
			en := readWQ(id)
			if (en.node="done") {
				return
			}
			parsePrevElement(id,en,res,"name")											; update elements if necessary
			parsePrevElement(id,en,res,"mrn")
			parsePrevElement(id,en,res,"date")
			parsePrevElement(id,en,res,"dev")
			parsePrevElement(id,en,res,"prov")
			parsePrevElement(id,en,res,"site")
			parsePrevElement(id,en,res,"duration")
			checkweb(id)
			return
		}
		if (id:=enrollcheck("[name=`"" res.name "`"]"									; 6/6 perfect match
			. "[mrn='" res.mrn "']"
			. "[date='" res.date "']"
			. "[dev='" res.dev "']"
			. "[prov=`"" res.prov "`"]"
			. "[site='" res.site "']" )) {
			en := readWQ(id)
			if (en.node="done") {
				return
			}
			parsePrevElement(id,en,res,"duration")
			checkweb(id)
			return
		}
		if (id:=enrollcheck("[name=`"" res.name "`"]"									; 4/6 perfect match
			. "[mrn='" res.mrn "']"														; everything but PROV or SITE
			. "[date='" res.date "']"
			. "[dev='" res.dev "']" )) {
			en:=readWQ(id)
			if (en.node="done") {
				return
			}
			eventlog("parsePrevEnroll " id "." en.node " changed PROV+SITE - matched NAME+MRN+DATE+DEV.")
			parsePrevElement(id,en,res,"prov")
			parsePrevElement(id,en,res,"site")
			parsePrevElement(id,en,res,"duration")
			checkweb(id)
			return
		}
		if (id:=enrollcheck("[mrn='" res.mrn "']"										; Probably perfect MRN+S/N+DATE
			. "[date='" res.date "']"
			. "[dev='" res.dev "']" )) {
			en:=readWQ(id)
			if (en.node="done") {
				return
			}
			eventlog("parsePrevEnroll " id "." en.node " changed NAME+PROV+SITE - matched MRN+DEV+DATE.")
			parsePrevElement(id,en,res,"name")
			parsePrevElement(id,en,res,"prov")
			parsePrevElement(id,en,res,"site")
			parsePrevElement(id,en,res,"duration")
			checkweb(id)
			return
		}
		if (id:=enrollcheck("[mrn='" res.mrn "'][date='" res.date "']")) {				; MRN+DATE, no S/N
			en:=readWQ(id)
			if (en.node="done") {
				return
			}
			if (en.node="orders") {														; falls through if not in <pending> or <done>
				addPrevEnroll(id,res)													; create a <pending> record
				wqSetVal(id,"name",en.name)												; copy remaining values from order (en)
				wqSetVal(id,"order",en.order)
				wqSetVal(id,"accession",en.accession)
				wqSetVal(id,"accountnum",en.acctnum)
				wqSetVal(id,"encnum",en.encnum)
				wqSetVal(id,"ind",en.ind)
				wq.removeNode("/root/orders/enroll[@id='" id "']")
				eventlog("addPrevEnroll moved Order ID " id " for " en.name " to Pending.")
				return
			}
			eventlog("parsePrevEnroll " id "." en.node " added DEV - only matched MRN+DATE.")
			parsePrevElement(id,en,res,"dev")
			parsePrevElement(id,en,res,"duration")
			checkweb(id)
			return
		}
		if (id:=enrollcheck("[date='" res.date "'][dev='" res.dev "']")) {				; DATE+S/N, no MRN
			en:=readWQ(id)
			if (en.node="done") {
				return
			}
			eventlog("parsePrevEnroll " id "." en.node " added MRN - only matched DATE+DEV.")
			parsePrevElement(id,en,res,"mrn")
			parsePrevElement(id,en,res,"duration")
			checkweb(id)
			return
		} 
		if (id:=enrollcheck("[mrn='" res.mrn "'][dev='" res.dev "']")) {				; MRN+S/N, no DATE match
			en:=readWQ(id)
			if (en.node="done") {
				return
			}
			dt0:= dateDiff(en.date,res.date,"Days")
			if abs(dt0) < 7 {															; res.date less than 5d from en.date
				parsePrevElement(id,en,res,"date")										; prob just needs a date adjustment
				parsePrevElement(id,en,res,"duration")
				eventlog("parsePrevEnroll " id "." en.node " adjusted date - only matched MRN+DEV.")
			}
			checkweb(id)
			return
		}
		if (id:=wq.selectSingleNode("/root/orders/enroll[mrn='" res.mrn "']").getAttribute("id")) {
			en:=readWQ(id)																; MRN found in Orders
			dt0:=dateDiff(en.date,res.date,"Days")
			
			if abs(dt0) < 7 {															; res.date less than 5d from en.date
				addPrevEnroll(id,res)													; create a <pending> record
				wqSetVal(id,"order",en.order)
				wqSetVal(id,"accession",en.accession)
				wqSetVal(id,"accountnum",en.acctnum)
				wqSetVal(id,"encnum",en.encnum)
				wqSetVal(id,"prov",en.provname)
				wqSetVal(id,"dev",res.dev)
				wqSetVal(id,"date",res.date)
				wqSetVal(id,"ind",en.ind)
				wq.removeNode("/root/orders/enroll[@id='" id "']")
				eventlog("addPrevEnroll order ID " id " for " en.name " " en.mrn " matched MRN only, moved to Pending.")
				return
			}
		}
		loop (allpend:=wq.selectNodes("/root/pending/enroll[mrn='" res.mrn "']")).Length
		{
			k := allpend.item(A_index-1)
			kser := k.selectSingleNode("dev").text
			kdev := strX(kser,"",0,1," - ",0,3) 
			rdev := strX(res.dev,"",0,1," - ",0,3)
			if !(kdev~=rdev) {															; rdev (from prev.txt) doesn't match kdev (from enroll)
				Continue
			}

			id := k.getAttribute("id")
			kdate := k.selectSingleNode("date").text
			dt := abs(res.date-kdate)
			if (dt>1) && (dt<5)															; if Preventice registration (res.date) off from 1-5 days
			{
				wqSetVal(id,"date",res.date)
				wqSetVal(id,"dev",res.dev)
				checkweb(id)
				eventlog("parsePrevEnroll " id "." en.node " changed DATE from " kdate " to " res.date ".")
				return
			}
		}																				; anything else is probably a new registration
		
	/*	No match (i.e. unique record)
	*	add new record to PENDING
	*/
		id := makeUID()
		addPrevEnroll(id,res)
		eventlog("Found novel web registration " res.mrn " " res.name " " res.date ". addPrevEnroll id=" id)
	
	return
}

enrollcheck(params) {
	global wq
	id := ""

	try en := wq.selectSingleNode("//enroll" params)
	try id := en.getAttribute("id")
	
; 	returns id if finds a match, else null
	return id																			
}

parsePrevDev(txt) {
/*	Add new dev to /root/inventory
 */
	global wq
	el := StrSplit(txt,"|")
	dev := el[2]
	ser := el[3]
	res := dev " - " ser

	if IsObject(wq.selectSingleNode("/root/inventory/dev[@ser='" ser "']")) {			; already exists in Inventory
		return
	}
	
	wq.addElement("/root/inventory","dev",{model:dev,ser:ser})
	;~ eventlog("Added new Inventory dev " ser)
	
	return
}

parsePrevElement(id,en,res,el) {
	/*	Update <enroll/el> node with value from result of Preventice txt parse
	
		id	= UID
		en	= enrollment node
		res	= result obj from Preventice txt
		el	= element to check
	*/
		global wq
		
		if (res.%el% == en.%el%) {														; Attr[el] is same in EN (wq) as RES (txt)
			return																		; don't do anything
		}
		if (en.%el%) and (res.%el%="") {												; Never overwrite a node with NULL
			return
		}
		
		wqSetVal(id,el,res.%el%)
		eventlog(en.name " (" id ") changed WQ " el " '" en.%el% "' ==> '" res.%el% "'")
		
		return
	}
	
addPrevEnroll(id,res) {
/*	Create <enroll id> based on res object
*/
	global wq
	
	newID := "/root/pending/enroll[@id='" id "']"
	wq.addElement("enroll","/root/pending",{id:id})
	wq.addElement("date",newID,res.date)
	wq.addElement("name",newID,res.name)
	wq.addElement("mrn",newID,res.mrn)
	wq.addElement("dev",newID,res.dev)
	wq.addElement("prov",newID,res.prov)
	wq.addElement("site",newID,res.site)
	wq.addElement("webgrab",newID,A_Now)
	
	return
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
#Include sift3.ahk
#Include Peep.v2.ahk																	; This is only for debugging