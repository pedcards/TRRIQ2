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
	pb := progressbar("Initialization...")
	progressbar(pb,20)
}

ExitApp

__GUI_elements:
progressbar(var1:="",var2:="w200 h20 cBlue vPercent") {
/*	Creates a minimal progress bar using Title and Params
	If first var is a pbar object, set the percentage
*/
	if IsObject(var1) {
		var1["Percent"].Value := var2
		return
	}
	pbar := Gui(,var1)
	pbar.Opt("+Border +AlwaysOnTop -SysMenu" 
			((var1="") ? " -Caption" : "")
	)
	pbar.AddProgress(var2)
	pbar.Show()
	return pbar
}