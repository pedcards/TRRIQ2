#Requires AutoHotkey v2

/*	Newer string trim from skan, can use to parse \<tags>\</tags>

	H = Haystack 
	C = Case sensitivity [0=none, 1=B sensitive, 2=E sensitive, 3=both]
	B = Begin match
	E = End match
	BO = Begin offset
	EO = End offet
	BI = Begin instance
	EI = End instance
	BT = Begin (un)trim [positive=trim, negative=untrim]
	ET = End (un)trim
*/
xStr(H, C:=0, B:="", E:="", BO:=1, EO:=0, BI:=1, EI:=1, BT:=0, ET:=0) {          
	Local L, LB, LE, P1, P2, Q, F:=0 ; xStr v0.97_dev by SKAN on D1AL/D343 @ tiny.cc/xstr  
	
	P1 := ( L := StrLen(H) ) 
			  ? ( LB := StrLen(B) )
						? ( F := InStr(H, B, C&1, BO, BI) ) 
							 ? F+(BT="" ? LB : BT) 
							 : 0 
			  : ( Q := (BO=1 && BT>0 ? BT+1 : BO>0 ? BO : L+BO) )>1 ? Q : 1      
		  : 0 
	
	
	P2 := P1              
			  ?  ( LE := StrLen(E) ) 
						? ( F := InStr(H, E, C>>1, EO=0 ? (F ? F+LB : P1) : EO, EI) )   
							 ? F+LE-(ET=0 ? LE : ET) 
							 : 0 
			  : EO=0 ? (ET>0 ? L-ET+1 : L+1) : P1+EO  
		  : 0
	
	Return SubStr(H, !( ErrorLevel := !((P1) && (P2)>=P1) ) ? P1 : L+1, ( BO := Min(P2, L+1) )-P1)  
	}

/*	Search between two strings using RegEx terms 

	h = Haystack
	BS = beginning string
	BO = beginning offset
	BT = beginning trim, TRUE or FALSE
	ES = ending string
	ET = ending trim, TRUE or FALSE
	N = variable for next offset
*/
stRegX(h,BS:="",BO:=1,BT:=0, ES:="",ET:=0, &N:="") {
	rem:="[PimsxADJUXPSC(\`n)(\`r)(\`a)]+\)"
	pos0 := RegExMatch(h, BS~=rem ? "im" BS : "im)" BS, &bPat, BO<1 ? 1 : BO)
	pos1 := RegExMatch(h, ES~=rem ? "im" ES : "im)" ES, &ePat, pos0+bPat.len)
	N := pos1+((ET) ? 0 : ePat.len)
	return substr(h,pos0+((BT) ? bPat.len : 0), N-pos0-bPat.len)
}
		
/*	StrX for V2 (based on the original from Skan www.autohotkey.com/forum/topic51354.html)

	H = HayStack. The "Source Text"
	BS = BeginStr. 
		Pass a String that will result at the left extreme of Resultant String.
	BO = BeginOffset. 
		Number of Characters to omit from the left extreme of "Source Text" while searching for BeginStr
		Pass a 0 to search in reverse ( from right-to-left ) in "Source Text"
		If you intend to call StrX() from a Loop, pass the same variable used as 8th Parameter, which will simplify the parsing process.
	BT = BeginTrim.
		Number of characters to trim on the left extreme of Resultant String
		Pass the String length of BeginStr if you want to omit it from Resultant String
		Pass a Negative value if you want to expand the left extreme of Resultant String
	ES = EndStr. 
		Pass a String that will result at the right extreme of Resultant String
	EO = EndOffset. Can be only True or False.
		If False, EndStr will be searched from the end of Source Text.
		If True, search will be conducted from the search result offset of BeginStr or from offset 1 whichever is applicable.
	ET = EndTrim.
		Number of characters to trim on the right extreme of Resultant String
		Pass the String length of EndStr if you want to omit it from Resultant String
		Pass a Negative value if you want to expand the right extreme of Resultant String
	NextOffset.
		A name of ByRef Variable that will be updated by StrX() with the current offset
		You may pass the same variable as Parameter 3, to simplify data parsing in a loop
*/
StrX(H,  BS:="",BO:=0,BT:=1,  ES:="",EO:=0,ET:=1,  &N:="") {
	X := StrLen(H)
	Y := StrLen(BS)
	Z := StrLen(ES)

	P1 := (Y)																			; BO=0 reverse searches from end 
		? InStr(H,BS,0,((BO)?BO:-1))													; Y>0, search from BO
		: 1																				; Y=0, start from 1

	if (EO) {																			; e0=1, search for es beginning after BS
		P2 := (Z)
			? InStr(H,ES,0,P1+Y)														; Z>0, begin search at P1+Y
			: X+1																		; Z=0, return end

	} else {																			; e0=0, reverse search for es from x 
		P2 := (Z)
			? (																			; Z>0, search for es from end
				((T:=InStr(H,ES,0,-1))>P1)
					? T																	; P2>P1 returns P2
					: 1																	; P2<=P1, returns 1
				)
			: X																			; Z=0, return end
	}

	N := P2+Z-ET

	return SubStr(H,P1+BT,(P2+Z)-(P1+BT)-ET)
}