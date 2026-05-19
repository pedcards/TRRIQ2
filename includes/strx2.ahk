#Requires AutoHotkey v2

/*	Newer string trim from skan, can use to parse \<tags>\</tags>
*/
xStr( H_, F   := 0                 ;  Haystack, Flags (Case sensitivity)  ;           xStr v1.00 by SKAN
        , B   := "",   E  := ""    ;  Begin match, End match              ;         for ah2 on D1AL/D92J
        , BO_ := 1,    EO := ""    ;  Begin offset, End offset            ;  @ autohotkey.com/r?t=140202
        , BI  := 1,    EI := 1     ;  Begin instance, End instance
        , BT  := "",   ET := "" )  ;  Begin (un)trim, End (un)trim
{
    Local  P  := 0,   L, LB, LE, P1, P2, Q
        ,  H  := (H_  is VarRef ? H_  : &H_)
        ,  BO := (BO_ is VarRef ? BO_ : &BO_)        

    P1 := ( L := StrLen(%H%) )
            ? ( LB := StrLen(B) )
                ? ( P := InStr(%H%, B, F & 1, %BO%, BI) )
                    ? P + (BT = "" ? LB : BT)
                    : ( F >> 2 & 1 )
                : ( Q := (%BO% = 1 && BT != "" ? BT + 1 : %BO% > 0 ? %BO% : L + %BO%) ) > 1 ? Q : 1
            : 0

  , P2 := ( P1 )
            ? ( LE := StrLen(E) )
                ? ( P := InStr(%H%, E, F >> 1 & 1, EO = "" ? ( P ? P + LB : P1 ) : EO, EI) )
                    ? P + LE - (ET = "" ? LE : ET)
                    : ( F >> 3 & 1 ? L + 1 : 0 )
                : ( EO = "" ) ? (ET != "" ? L - ET + 1 : L + 1) : P1 + EO
            : 0

    Return SubStr( %H%, !(xStr.Error := !((P1) && (P2) >= P1)) ? P1 : L + 1
                      , (%BO% := Min(P2, L + 1)) - P1 )
}

/*	Wrapper function for xStr

	Str := "Item1|Item2|Item3|Item4|Item5|Item6|Item7|Item8|Item9"
	MsgBox GetStr(&Str)        ;  Item1
	MsgBox GetStr(&Str, 5)     ;  Item5
	MsgBox GetStr(&Str, 7, 2)  ;  Item7|Item8
	MsgBox GetStr(&Str, 10)    ;
*/
GetStr(Str, I := 1, C := 1, D := "|")  =>  xStr(Str, 0x8, I > 1 ? D : "", D,,, I - 1, C)

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
	=0)?(-1):(BO))))?(T+BT):(X+1)):(1)),(N:=P+((Z)?((T:=InStr(H,ES,0,((EO)?(P+1):(-1))))?(T-P+Z
	+(0-ET)):(X+P)):(X)))-P) ; v1.0-196c 21-Nov-2009 www.autohotkey.com/forum/topic51354.html
	}