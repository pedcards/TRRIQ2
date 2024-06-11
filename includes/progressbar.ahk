#Requires AutoHotkey v2
class progressbar
{
	; progressbar params (in any order):
	; "w400 h12","TITLE","subtitle"
	; Can only update elements added at creation
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

	set(val:=0) {
		if (val~="^[+-]") {
			try this.gui["Percent"].Value += val
		} else {
			try this.gui["Percent"].Value := val
		}
		try this.gui.Show()
	}
	title(val) {
		try this.gui["Title"].Value := val
		try this.gui.Show()
	}
	sub(val) {
		try this.gui["Subtitle"].Value := val
		try this.gui.Show()
	}
	hide() {
		try this.gui.Hide()
	}
	close() {
		try {
			this.gui.Destroy()
			this.gui := ""
		}
	}
}
