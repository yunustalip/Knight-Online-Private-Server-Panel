METÝN FONKSÝYONLARI
<% Metnimiz="Bugün VBScript öðreniyoruz"
   metintest="abcde ABCDE ABCDE ABCDE ABCDE" %> 
Uzunluk = <%=len(Metnimiz)%><br />
<%aranacakmetin="öð"%>
pozisyon = <%=Instr(Metnimiz,aranacakmetin)%><br />
terspozisyon = <%=InstrRev(metintest,"CDE")%><br />
terspozisyon = <%=InstrRev(metintest,"CDE",20)%><br />
terssirada = <%=StrReverse(Metnimiz)%><br />
terssirada = <%=StrReverse(Metintest)%><br />
buyuk = <%=Ucase(Metnimiz)%><br />
kucuk = <%=Lcase(Metnimiz)%><br />
bosluk = "<%="a"&space(15)&"z"%>"<br />
string = "<%=string(15,"-")%>"<br />
<%  metintest="     ABCDE ABCDE       "  %> 
Ltrim = "<%=Ltrim(Metintest)%>"<br />
Rtrim = "<%=Rtrim(Metintest)%>"<br />
Trim = "<%=Trim(Metintest)%>"<br />
replace = "<%=Replace(Ltrim(Metintest)," ","&nbsp;")%>"<br />
replace = "<%=Replace(Rtrim(Metintest)," ","&nbsp;")%>"<br />
sol= <%=Left(Metnimiz,10)%><br />
sag = <%=Right(Metnimiz,11)%><br />
sag = <%=Mid(Metnimiz,10,5)%><br />
<%metintest="abcde ABCDE ABCDE ABCDE ABCDE" %> 
replace = <%=Replace(metintest,"BCD","XYZ",7,2)%><br />
<% bosluk = "a"&space(15)&"z<br />"%>
replace = <%=Replace(bosluk," ","&nbsp;")%><br />
Asc = <%=asc("Z")%><br />
chr = <%=chr(90)%><br />
chr = <%=chr(169)%><br />
