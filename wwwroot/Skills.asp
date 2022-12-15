<!--#include file="_inc/conn.asp"-->
<!--#include file="function.asp"-->&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%Response.Expires=0
Dim skillbarno,sid,clss,sir
skillbarno=Request.Querystring("skillbarno")
sid=secur(Request.Querystring("sid"))
if skillbarno="" Then
skillbarno="1"
End If
skillbarno=cint(skillbarno)
function clss2(tur)
select case tur
case "101", "105", "106", "201", "205", "206"
clss2="Warrior"
case "102", "107", "108", "202", "207", "208"
clss2="Rogue"
case "103", "109", "110", "203", "209", "210"
clss2="Mage"
case "104", "111", "112", "204", "211", "212"
clss2="Priest"
Case else
clss2="Unknown"
end select
end function

set clss=Conne.Execute("select class from userdata where struserid='"&sid&"'")
if not clss.eof Then
Conne.Execute("exec skillbardecode '"&sid&"'")
dim skillnm,snm,skillnum,skillname,skiln,des,sklnm,slvl
for sir=1 to 8
skillnm="No Skill"
set snm=Conne.Execute("select * from skillbar where charid='"&sid&"' and sira="&sir&" and satir="&skillbarno&"")
if not snm.eof Then
skillnum=snm("skillno")

set skillname=Conne.Execute("select enname,krname,description from magic where magicnum="&skillnum)
if not skillname.eof Then
skiln=mid(skillnum,4,1)
des=skillname("description")
if instr(skillname(1),"?")>0 Then
skillnm=skillname(0)
else
skillnm=skillname(1)
End If

End If

if clss2(clss(0))="Warrior" Then
if skiln=5 Then
sklnm="Attack"
elseif skiln=6 Then
sklnm="Defense"
elseif skiln=7 Then
sklnm="Berserker"
elseif skiln=8 Then
sklnm="Master"
End If
End If

if clss2(clss(0))="Mage" Then
if skiln=5 Then
sklnm = "Flame"
elseif skiln=6 Then
sklnm = "Glacier"
elseif skiln=7 Then
sklnm = "Lightning"
elseif skiln=8 Then
sklnm="Master"
End If
End If

if clss2(clss(0))="Priest" Then
if skiln=5 Then
sklnm = "Heal"
elseif skiln=6 Then
sklnm = "Buff"
elseif skiln=7 Then
sklnm = "Debuff"
elseif skiln=8 Then
sklnm="Master"
End If
End If

if clss2(clss(0))="Rogue" Then
if skiln=5 Then
sklnm = "Archer"
elseif skiln=6 Then
sklnm = "Assassin"
elseif skiln=7 Then
sklnm = "Explore"
elseif skiln=8 Then
sklnm="Master"
End If
End If

slvl="("&sklnm&" "&mid(skillnum,5,2)&")"

if mid(skillnum,4,1)="0" Then
slvl="(item)"
End If

Response.Write "<span style=""position:relative;margin-left:2px;z-index:3""  onMouseOver=""return overlib('<body bgcolor=#000000><b><center><font style=font-size:11px color=white>"&server.htmlencode(skillnm)&slvl&"<br><br>"&des&"', ABOVE, WIDTH, 300,CELLPAD, 5, 5, 5);"" onMouseOut=""return nd();""><img src=""skill/skillicon_"&mid(skillnum,5,2)&"_"&mid(skillnum,1,4)&".bmp""></span>"
else
Response.Write "<span style=""position:relative;margin-left:2px;z-index:3"" onMouseOver=""return overlib('<body bgcolor=#000000><b><center><font style=font-size:11px color=white>No Skill', ABOVE, WIDTH, 200,CELLPAD, 0, 0, 5);"" onMouseOut=""return nd();"")""><img src=""skill/skillicon_enigma.bmp""></span>"
End If
des=""
Next
Conne.Execute("Delete skillbar where charid='"&sid&"'")
End If
%><img src="../imgs/skillbar.gif"  style="position:relative; left:-303px; top: 3px; z-index:1">
<div style="position:relative;left: 13px; top:-33px;width:10px; z-index:3"><%if skillbarno>1 Then
Response.Write "<a onclick=""loadskill("&skillbarno-1&")"">"
End If%><img src="imgs/yukok.gif" border="0" width="9" height="9" style="position:relative;z-index:3"><br><font style="color:#FFFFFF;font-size:10px;position:relative; z-index:3"><b><%=skillbarno%></b></font><%if skillbarno<8 Then
Response.Write"<a onclick=""loadskill("&skillbarno+1&")"">"
End If%><br><img src="imgs/asok.gif" border="0" width="9" height="9" style="position: relative; top: -0px; width: 9px; height: 10px; z-index: 10;"></a><br />
<img src="imgs/skillbartop.gif" style="position:relative;top:-33px;left:13px;z-index:5">
</div>
