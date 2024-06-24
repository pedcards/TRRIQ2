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
		
/*	StrX for V2

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
StrX( H,  BS:="",BO:=0,BT:=1,   ES:="",EO:=0,ET:=1,  &N:="" ) { ;    | by Skan | 19-Nov-2009
	Return SubStr(H,P:=(((Z:=StrLen(ES))+(X:=StrLen(H))+StrLen(BS)-Z-X)?((T:=InStr(H,BS,0,((BO
	=0)?(-1):(BO))))?(T+BT):(X+1)):(1)),(N:=P+((Z)?((T:=InStr(H,ES,0,((EO)?(P+1):(0))))?(T-P+Z
	+(0-ET)):(X+P)):(X)))-P) ; v1.0-196c 21-Nov-2009 www.autohotkey.com/forum/topic51354.html
	}