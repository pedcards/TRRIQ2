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