#Requires AutoHotkey v2
today := FormatTime(A_Now,"Wday")
last := DateAdd(A_Now,2-today,"Days")
