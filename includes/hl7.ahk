#Requires AutoHotkey v2

initHL7() {
	global hl7ref, prevDDE

	inifile := ".\files\hl7.ini"

	hl7ref := Map()
	s0 := IniRead(inifile)																; s0 = Section headers
	loop parse s0, "`n", "`r"															; parse s0
	{
		i := A_LoopField
		hl7ref.%i% := Map()																; create array for each header
		s1 := IniRead(inifile, i)														; s1 = individual header
		loop parse s1, "`n", "`r"														; parse s1
		{
			j := A_LoopField
			arr := strSplit(j,"=",,2)													; split into arr.1 and arr.2
			num := arr[1]+0
			hl7ref.%i%.%num% := arr[2]													; set hl7.OBX.2 = "Obs Type"
		}
	}
	prevDDE := readIni("preventiceDDE")													; map hl7 fields to lw fields
}

class hl7
{
	__New(fnam:="") {
	/*	Initialize this.map
	 */
		if (fnam="") {
			return
		}
		this.file := FileRead(fnam)														; store the name of this HL7 file	
		this.fldval := Map()															; values from segment fields
		this.fldval["hl7string"] := ""													; to store cumulative string (ultimately same as input file)
		this.obxval := Map()															; result values
		this.bin := ""																	; extracted binary
		this.binfile := ""																; binary filename
		this.hl7out := ""																; generated HL7 segment

		this.processHL7(fnam)
	}
	
	processHL7(fnam) {
	/*	Read an HL7 file and parse each line
	*/
		txt := FileRead(fnam)
		txt := StrReplace(txt, "`r`n", "`r")											; convert `r`n to `r
		txt := StrReplace(txt, "`n", "`r")												; convert `n to `r
		loop parse txt, "`r", "`n"														; parse HL7 message, split on `r, ignore `n
		{
			seg := A_LoopField															; read next Segment line
			if (seg=="") {
				continue
			}
			this.hl7line(seg)
		}
	}

	hl7line(seg) {
	/*	Interpret an hl7 message "segment" (line)
		segments are comprised of fields separated by "|" char
		field elements can contain subelements separated by "^" char
		field elements stored in res[i] object
		attempt to map each field to recognized structure for that field element
	*/
		global path, hl7ref, prevDDE
		multiSeg := "NK1|DG1|NTE"														; segments that may have multiple lines, e.g. NK1
		
		fld := StrSplit(seg,"|")														; split on `|` field separator into fld array
		segName := fld[1]																; first array element should be NAME
		segNum := fld[2]
		if !IsObject(hl7ref.%segName%) {												; no matching segment in hl7 reference?
			MsgBox(seg "-" segName "`nBAD SEGMENT NAME",A_Index)
			return error																; fail if segment name not allowed
		}

		isOBX := (segName == "OBX")
		segMap := hl7ref.%segName%														; get the reference map for this segment name
		segPre := segName . (instr(multiSeg,segName) ? "_" segNum : "")					; number PID_1 PID_2 PID_3 if multiple
		this.fldval[segPre] := Map()													; create value map() for each found segment

		res := Map()																	; values for each field/subfield
		Loop fld.Length																	; step through each of the fld[] strings
		{
			i := A_Index
			if (i<=1) {																	; skip first 2 elements in OBX|2|TX
				continue
			}
			str := fld[i]																; each segment field value, e.g. PID.3(MRN)=1494708
			val := StrSplit(str,"^")													; array of subelements
			if (str="") {
				val := [""]
			}
			this.fldval[segPre][i-1] := str												; this.fldval["PID"][2]=1494708

			try {
				strMap := segMap.%i-1%													; get hl7 field name that maps to this
			}
			catch {
				strMap := ""
			}
			if (strMap=="") {															; no mapped fields
				loop val.length															; create strMap "^^^" based on subelements in val
				{
					strMap .= "z" i "_" A_Index "^"
				}
			}

			submap := StrSplit(strMap,"^")												; array of substring map
			loop submap.length
			{
				j := A_Index
				if (submap[j]=="") {													; skip if map value is null
					continue
				}
				x := strQ(segPre,"###_") submap[j]										; res.pre_map

				if (submap.length=1) {													; for seg with only 1 map, ensure val is at least popuated with str
					val[j] := str
				}
				if (j>val.length) {
					val.length := j
					val[j] := ""
				}
				res[x] := val[j]														; add each mapped result as subelement, res.mapped_name

				if !(isOBX)  {															; non-OBX results
					this.fldval[x] := val[j]											; populate all fldVal.mapped_name
				}
			}
		}
		if (isOBX) {																	; need to special process OBX[], test result strings
			if (res.ObsType == "ED") {
				this.binfile := res.Filename											; file follows
				b64 := res.resValue
				bin := this.Base64_Dec(&b64)
				Fx := FileOpen( path.PrevHL7in . res.Filename, "w")
				Fx.RawWrite(bin)
				Fx.Close()
			} else {
				label := res.resCode													; result value
				result := strQ(res.resValue, "###")
				maplab := strQ(this.DDE[label],"###",label)								; maps label if hl7->lw map exists
						. strQ(res.Filename,"_###")        								; add suffix if multiple units in OBX_Filename
				this.fldval[segPre maplab] := result
				this.obxval[segPre maplab] := result
			}
		}
		this.fldval["hl7string"] .= seg "`n"

		return res
	}

	Base64_Dec(&Src)                                                        ;  By SKAN for ah2 on D672/D672 @ autohotkey.com/r?p=534720
	{
		Local  EqTo    :=  (SubStr(Src,-2,1) = "=") + (SubStr(Src,-1) = "=")   ;  = count
			,  nBytes  :=  (StrLen(Src) - EqTo) * 3 // 4                         ;  Target bytes
			,  Trg     :=  Buffer(nBytes)

		DllCall("Crypt32\CryptStringToBinary", "str",Src, "int",StrLen(Src), "int",0x1, "ptr",Trg, "intp",&nBytes, "int",0, "int",0 )

	Return Trg
	}
	Base64_Enc(&Src)                                                        ;  By SKAN for ah2 on D672/D672 @ autohotkey.com/r?p=534720
	{
		Local  Bytes  :=  Src.Size
			,  RqdCap :=  1 + (( Ceil(Bytes*4/3) + 3 ) & ~0x03)
			,  Trg    :=  ""

		VarSetStrCapacity(&Trg, RqdCap - 1)
		DllCall("Crypt32\CryptBinaryToString", "ptr",Src, "int",Bytes, "int",0x40000001, "str",Trg, "intp",&RqdCap)

	Return Trg
	}

	segField(fld,lbl:="") {
		res := XML("<root/>")
		split := StrSplit(fld,"~")
		loop split.length()
		{
			i := A_index
			res.addElement("/root","idx",{num:i})
			id := "/root/idx[@num='" i "']"
			subfld := StrSplit(split[i],"^")
			sublbl := StrSplit(lbl,"^")
			loop subfld.length()
			{
				j := A_Index
				k := sublbl[j]
				if (k="") {
					res.addElement(id,"node",{num:j},subfld[j])
				}
				else {
					res.addElement(id,k,subfld[j])
				}
			}
		}
		return res
	}

	buildHL7(seg,params) {
	/*	creates hl7out.msg = "seg|idx|param1|param2|param3|param4|..."
		keeps seg counts in hl7out[seg] = idx
		params is a sparse object with {2:"TX", 3:str1, 5:value, 11:"F", 14:A_now}
	*/
		txt := seg
	
		Loop params.MaxIndex()
		{
			param := params[A_index]
	
			if (seg!="MSH")&&(A_index=1) {
				seqnum := this.hl7out[seg]														; get last sequence number for this segment
				++ seqnum
				this.hl7out[seg] := seqnum
				param := seqnum
			}
	
			txt .= "|" param
	
		}
	
		this.hl7out.msg .= txt "`n"																; append result to hl7out.msg
	
		return
	}
}
