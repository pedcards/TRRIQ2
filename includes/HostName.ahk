#Requires AutoHotkey v2
;****************************************************************************************
; Module: HostName
; Purpose: Determine which clinic an SCH workstation/vdi is installed in.  If the script
;          does not find the workstation/vdi in the xml file (e.g. first time the workstation or
;          vdi has run the script) the script will prompt the user for their location and persist
;          the location to the xml file
;
; Assumptions:
;    -- In general, workstations do not change locations
;    -- Location data for the xml file is up to date (includes all valid locations)
;    -- XML hierarchy: root -> locations
;                      root -> workstations -> workstation -> wsname
;                      root -> workstations -> workstation -> location
;
class getLocation
{
	; **** Globals used for GUI
	SelectedLocation := ""      
	SelectConfirm := ""

	; **** Globals used as constants (do not change these variables in the code)
	m_strXmlFilename := ".\data\wkslocation.xml"                          ; path to xml data file that contains workstation information
	m_strXmlLocationsPath := "/root/locations"                            ; xml path to locations node (location names)
	m_strXmlWorkstationsPath := "/root/workstations"                      ; xml path to workstations node (contains all infomation for workstations)
	m_strXmlWksNodeName := "workstation"                                  ; name of "workstation" node in the xml data file
	m_strXmlWksName := "wsname"                                           ; name of the "workstation name node" in the xml data file
	m_strXmlLocationName := "location"                                    ; name of the "location" node in the xml data file

	__New() {
		wks := A_ComputerName
		location := this.GetWksLocation(wks)
		this.location := location
	}

	;******************************************************************************
	; Function: GetWksLocation
	; Purpose : Retrieve the location for the specified workstation from the xml
	;           file
	; Output  : Success = string containing workstation's name
	;           Falure = empty string
	; Input   : nameIn - string containing the hostname for the current workstation
	; Assumptions : 
	;     - xml file named wkslocation.xml
	;     - xml file in same folder as script
	;     - xml hierarchy is known and static
	;     - if workstation is not found user should be prompted for location
	;     - return empty string on failure
	;
	GetWksLocation(nameIn)
	{
		location := ""																		; assume failure

		if FileExist(this.m_strXmlFilename) {
			locationData := XML(this.m_strXmlFilename)										; load xml file
			wksList := locationData.SelectSingleNode(this.m_strXmlWorkstationsPath)			; retreive list of all workstations
			wksFound := false
			loop (wsNodes := wksList.SelectNodes(this.m_strXmlWksNodeName)).Length {		; loop through the workstations
				wsInfoNode := wsNodes.item(A_Index - 1)										; Retrieve workstation node from workstation list
				wsName     := wsInfoNode.SelectNodes(this.m_strXmlWksName).item(0).Text		; Retrieve the wsname node from the workstation informaton
					
				if (wsName = nameIn)														; compare this workstation to current workstation list name
				{
					location := wsInfoNode.SelectNodes(this.m_strXmlLocationName).item(0).Text	; names matched, retreive the workstation location
					wksfound := true
					break
				}
			}
		
			if (wksfound = false) {
				location := this.PromptForLocation()										; Prompt user for location of new workstation
			}
		
		} else {
				MsgBox(
					"Location data unavailable: The file " this.m_strXmlFilename " was not found.",
					"File Error",
					16
				)
		}
		return location
	}

	;******************************************************************************
	; Function: PromptForLocation
	; Purpose : Retrieve the location of the current workstation from the user by
	;           displaying a dialog that will allow the user to select their loctaion 
	;           from a list of all the available locations.  After user selects the location
	;           call the function to persist the location to the data store.
	;
	PromptForLocation()
	{
		workstationLocation := ""
		SelectConfirm := ""
		SelectedLocation := ""

		locationData := this.GetLocations()                                              ; Function to retrive the location list from the data store

		;Buld and display the dialog box

		wksGUI := Gui()
		wksGUI.Title := "Unknown Location"
		wksGUI.AddText("x15 y20 w250 h60","The application is unable to determine your location. Please select your location from the list and confirm that you made the correct selection.")
		wksGUI.AddListBox("vSelectedLocation x15 y70 w245 h200 Sort gLocationList_Click",locationData)
		wksGUI.AddText("x100 y290 w72","You selected:")
		wksGUI.AddText("vSelectConfirm x172 y290 w150",SelectConfirm)
		wksGUI.AddButton("x160 y315 w100 gConfirmBtn_Click","Confirm")
		wksGUI.Opt("AlwaysOnTop -MaximizeBox -MinimizeBox")
		wksGUI.Show()


		; Gui, New, AlwaysOnTop -MaximizeBox -MinimizeBox, Unknown Location
		; Gui, Add, Text, x15 y20 w250 h60,The application is unable to determine your location. Please select your location from the list and confirm that you made the correct selection.
		; Gui, Add, ListBox, vSelectedLocation x15 y70 w245 h200 Sort gLocationList_Click, %locationData%
		; Gui, Add, Text,x100 y290 w72, You selected:
		; Gui, Add, Text, vSelectConfirm x172 y290 w150, %SelectConfirm% 
		; Gui, Add, Button, x160 y315 w100 gConfirmBtn_Click, Confirm
		; Gui, Show, w275 h350
		
		WinWaitClose("Unknown Location")                                              ;wait for the user to respond
		return %workstationLocation%                                                ;return the selected location

		;******************* Gui Event handlers (subroutines) *********************
		LocationList_Click:
			wksGUI.Submit([0])                                                     ; user selected location from list, submit dialog data / keep displaying the dialog
			wksGUI.Control()
			wksGUI.Text := SelectedLocation                          ; reflect selected value in confirmation text box
		return

		ConfirmBtn_Click:
			wksGUI.Submit([1])
			this.AddWorkstation(SelectedLocation)                                        ; Persist workstation/location to data store
			workstationLocation := SelectedLocation                                 ; set the return value to the selected location
			WinClose("Unknown Location")                                            ; Close the dialog
			wksGUI.Destroy()                                                        ; Release resources
		return
	}

	;******************************************************************************
	; Function: GetLocations
	; Purpose : Retrieve the location list from the data store in a format compatible
	;           for use in a Gui ListBox control
	; Output  : String contining piped list of locations
	; Input   : N/A
	;
	GetLocations()
	{
		locationList := ""
		
		locationData := XML(this.m_strXmlFilename)                       ; Read xml file
		
		wksList := locationData.SelectSingleNode(this.m_strXmlLocationsPath)      ; Retreive Locations node
		loop (wksNodes := wksList.SelectNodes(this.m_strXmlLocationName)).Length     ; Loop through node and create piped list of locations
		{
			location:= wksNodes.item(A_Index - 1).selectSingleNode("site").text
			if (A_Index = 1) {
				locationList := location                                 ; No pipe symbol before fist location
			} else {
				locationList := locationList . "|" . location
			}
		}
		return locationList
	}

	;******************************************************************************
	; Function: AddWorkstation
	; Purpose : Persist the workstation/location to the data store
	; Output  : N/A
	; Input   : locationData - pointer to the 
	;
	AddWorkstation(location)
	{
		global gl

		if (ObjHasValue(gl.wksVoid,A_ComputerName,1)) {								; don't write if in wksVM list
			Return
		}

		locationData := XML(this.m_strXmlFilename) 
		
		workstations := locationData.SelectSingleNode(this.m_strXmlWorkstationsPath)
		workstation := locationData.addElement(workstations,this.m_strXmlWksNodeName,A_ComputerName)
		locationData.addElement(workstation,this.m_strXmlLocationName,location)
		locationData.saveXML()
		; workstation := locationData.addChild(this.m_strXmlWorkstationsPath,
		; 							"element",
		; 							this.m_strXmlWksNodeName)
		
		; wsnameNode := locationData.createNode(1,this.m_strXmlWksName,"")
		; wsnameNode.Text := A_ComputerName
		; workstation.appendChild(wsnameNode)
		
		; locationNode := locationData.createNode(1,this.m_strXmlLocationName,"")
		; locationNode.Text := location
		; workstation.appendChild(locationNode)
		
		; locationData.TransformXML()
		; locationData.saveXML()
		
		eventlog("New machine " workstation.Text " assigned to location " location)
	}

	getSites(wksName) {
	/*	reads wkslocation.xls and returns:
			sites (MAIN|BELLEVUE|EVERETT...) and sites0 (TACOMA|ALASKA...) menus
			sitesLong {EKG:MAIN,INPATIENT:MAIN,CRDBCSC:BELLEVUE,...}
			facility code {MAIN:7343,...}
			facility name {MAIN:GB-SCH-MAIN,...}
		wksName argument returns the hl7code and hl7name (Preventice facility codes)
	*/
		locationList := Map(0,"",1,"")
		locationLong := Map()
		locationData := XML(this.m_strXmlFilename)
		wksList := locationData.SelectSingleNode(this.m_strXmlLocationsPath)
		loop (wksNodes := wksList.SelectNodes(this.m_strXmlLocationName)).Length
		{
			tracked := 1
			location:= wksNodes.item(A_Index - 1)
			try tabname := location.selectSingleNode("tabname").text
			try tracked := !(location.selectSingleNode("tracked").text = "n")
			locationList[tracked] .= tabname . "|"
		}
		loop (wksNodes := wksList.SelectNodes(this.m_strXmlLocationName "/alias")).Length
		{
			node := wksNodes.item(A_Index-1)
			aliasName := node.text
			longName := node.selectSingleNode("../tabname").text
			locationLong[aliasName] := longName
		}
		wksNode := wksList.selectSingleNode(this.m_strXmlLocationName "[site='" wksName "']")
		codeName := wksNode.selectSingleNode("hl7name").text
		codeNum := wksNode.selectSingleNode("hl7num").text
		tabname := wksnode.selectSingleNode("tabname").text
		
		return {  tracked:trim(locationList[1],"|")
				, ignored:trim(locationList[0],"|")
				, long:locationLong
				, code:codeNum
				, facility:codeName
				, tab:tabname}
	}	
}
#Include xml2.ahk