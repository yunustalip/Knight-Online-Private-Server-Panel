<!--#include file="_inc/conn.asp"-->
<!--#include file="function.asp"-->
<%classx=Request.Querystring("class")
nations=Request.Querystring("nation")
skilltur=Request.Querystring("skilltur")
table=Request.Querystring("table")



if table="" Then
table="magic"
End If%><style>
a{
color:#000;
text-decoration:none
}
body {
	font-family: "Lucida Grande", Verdana, sans-serif;
	background-color: none;
	font-size: 8px;
	line-height: 1.7em;
	padding: 0;
	margin: 0;
	text-align: center;
}
td{
font-size:12px
}
.tdx{
font-size:11px;
font-weight:bold
}
.tvx{
height:18px;
font-size:10px;
font-weight:bold
}
</style>
<form action="?" method="get">
<select name="class" onchange="this.form.submit();this.skilltur.value=7">
<option value="05" <%if classx="05" Then Response.Write "selected"%>>Warrior</option>
<option value="06" <%if classx="06" Then Response.Write "selected"%>>Master Warrior</option>
<option value="07" <%if classx="07" Then Response.Write "selected"%>>Rogue</option>
<option value="08" <%if classx="08" Then Response.Write "selected"%>>Master Rogue</option>
<option value="11" <%if classx="11" Then Response.Write "selected"%>>Priest</option>
<option value="12" <%if classx="12" Then Response.Write "selected"%>>Master Priest</option>
<option value="09" <%if classx="09" Then Response.Write "selected"%>>Mage</option>
<option value="10" <%if classx="10" Then Response.Write "selected"%>>Master Mage</option>
</select>
<select name="nation" onchange="this.form.submit()">
<option value="2" <%if nations="2" Then Response.Write "selected"%>>Human</option>
<option value="1" <%if nations="1" Then Response.Write "selected"%>>Karus</option>
</select><select name="skilltur" onchange="this.form.submit()" id="skilltur">
<option value="0">Basic Skills</option>
<%if classx="" Then
classx="05"
End If
if nations="" Then
nations="2"
End If
if classx="05" and skilltur="8" or classx="07" and skilltur="8" or classx="09" and skilltur="8" or classx="11" and skilltur="8" Then
skilltur="0"
End If

if skilltur="" Then skilltur="0"
if classx="05" Then
Response.Write "<option value=""5"" "
if skilltur=5 Then Response.Write "selected"
Response.Write ">Attack</option>"
Response.Write "<option value=""6"" "
if skilltur=6 Then Response.Write "selected"
Response.Write ">Defense</option>"
Response.Write "<option value=""7"" "
if skilltur=7 Then Response.Write "selected"
Response.Write ">Berserker</option>"
elseif classx="06" Then
Response.Write "<option value=""5"" "
if skilltur=5 Then Response.Write "selected"
Response.Write ">Attack</option>"
Response.Write "<option value=""6"" "
if skilltur=6 Then Response.Write "selected"
Response.Write ">Defense</option>"
Response.Write "<option value=""7"" "
if skilltur=7 Then Response.Write "selected"
Response.Write ">Berserker</option>"
Response.Write "<option value=""8"" "
if skilltur=8 Then Response.Write "selected"
Response.Write ">Master</option></option>"
End If

if classx="09" Then
Response.Write "<option value=""5"" "
if skilltur=5 Then Response.Write "selected"
Response.Write ">Flame</option>"
Response.Write "<option value=""6"" "
if skilltur=6   Then Response.Write "selected"
Response.Write ">Glacier</option>"
Response.Write "<option value=""7"" "
if skilltur=7  Then Response.Write "selected"
Response.Write ">Lightning</option>"

elseif classx="10" Then
Response.Write "<option value=""5"" "
if skilltur=5  Then Response.Write "selected"
Response.Write ">Flame</option>"
Response.Write "<option value=""6"" "
if skilltur=6 Then Response.Write "selected"
Response.Write ">Glacier</option>"
Response.Write "<option value=""7"" "
if skilltur=7 Then Response.Write "selected"
Response.Write ">Lightning</option>"
Response.Write "<option value=""8"" "
if skilltur=8 Then Response.Write "selected"
Response.Write ">Master</option>"
End If

if classx="11" Then
Response.Write "<option value=""5"" "
if skilltur=5 Then Response.Write "selected"
Response.Write ">Heal</option>"
Response.Write "<option value=""6"" "
if skilltur=6 Then Response.Write "selected"
Response.Write ">Aura(Buff)</option>"
Response.Write "<option value=""7"" "
if skilltur=7 Then Response.Write "selected"
Response.Write ">Talisman(Debuff)</option>"
elseif classx="12" Then
Response.Write "<option value=""5"" "
if skilltur=5 Then Response.Write "selected"
Response.Write ">Heal</option>"
Response.Write "<option value=""6"" "
if skilltur=6 Then Response.Write "selected"
Response.Write ">Aura(Buff)</option>"
Response.Write "<option value=""7"" "
if skilltur=7 Then Response.Write "selected"
Response.Write ">Talisman(Debuff)</option>"
Response.Write "<option value=""8"" "
if skilltur=8 Then Response.Write "selected"
Response.Write ">Master</option>"
End If

if classx="07" Then
Response.Write "<option value=""5"" "
if skilltur=5 Then Response.Write "selected"
Response.Write ">Archer</option>"

Response.Write "<option value=""6"" "
if skilltur=6 Then Response.Write "selected"
Response.Write ">Assassin</option>"

Response.Write "<option value=""7"" "
if skilltur=7 Then Response.Write "selected"
Response.Write ">Explore</option>"

elseif classx="08" Then
Response.Write "<option value=""5"" "
if skilltur=5 Then Response.Write "selected"
Response.Write ">Archer</option>"

Response.Write "<option value=""6"" "
if skilltur=6 Then Response.Write "selected"
Response.Write ">Assassin</option>"

Response.Write "<option value=""7"" "
if skilltur=7 Then Response.Write "selected"
Response.Write ">Explore</option>"

Response.Write "<option value=""8"" "
if skilltur=8 Then Response.Write "selected"
Response.Write ">Master</option>"
End If
%></select>
<input type="submit" value="Ara"></form><table ><tr><td width="400" valign="top">
<div style="overflow:auto; height:500px;width:250px; border-right:black 2px solid;border-top:black 2px solid;scrollbar-highlight-color:silver;width:320px; background:#CCCCCC;border-bottom:black 2px solid;">
<%if Request.Querystring("islem")="kaydet" Then
Set skillupdate = Server.CreateObject("ADODB.RecordSet")
If table="magic" Then
SQL = "Select * from magic where MagicNum="&Request.Querystring("skillno")
Else
SQL = "Select * from "&table&" where inum="&Request.Querystring("skillno")
End If
Skillupdate.Open SQL,conne,1,3
For Each xy In Request.Form
skillupdate(cint(xy))=request.form(xy)
Next
Skillupdate.Update
Response.Redirect("skilleditor.asp?class="&classx&"&nation="&nations&"&skilltur="&skilltur)
End If

Set skill=Conne.Execute("Select * From Magic where MagicNum>="&nations&classx&skilltur&"00 and MagicNum<"&nations&classx&skilltur+1&"00  order by magicnum")
Response.Write "<table width=""300"" style=""font-weight:bold;font-size:13px;font-family: Lucida Grande;"">"

If Not skill.eof Then
Do While Not skill.eof
Response.Write "<tr id="""&skill(0)&""" onMouseOver=""this.style.background='#D5AB4A'"" onMouseOut=""this.style.background=''""><td><a href=""?nation="&nations&"&class="&classx&"&skilltur="&skilltur&"&skillno="&skill(0)&"#"&skill(0)&""" style=""display:block""><img border=""0"" align=""absmiddle"" src=""skill/skillicon_"&mid(skill(0),5,2)&"_"&mid(skill(0),1,4)&".bmp"" onError=src=""skill/skillicon_enigma.bmp""> "&vbcrlf
If instr(skill(2),"?")>0 Then
Response.Write trim(skill(1))
Else
Response.Write trim(skill(2))
End If
Response.Write " "&skill(0)&" </a></td></tr>"&vbcrlf
skill.MoveNext
Loop
End If
Response.Write "<script>document.getElementById('"&Request.Querystring("skillno")&"').style.background='#D5AB4A';</script>"

%></table></div>
</td>
<td valign="top"><table ><form method="post" action="?islem=kaydet&table=<%=table%>&skillno=<%=Request.Querystring("skillno")%>&nation=<%=nations%>&class=<%=classx%>&skilltur=<%=skilltur%>">
<%If Not Request.Querystring("skillno")="" Then
tablename=Request.Querystring("table")
If tablename="" Then
tablename="magic"
End If
Set skillcolumn=Conne.Execute("Select Name From syscolumns Where Id=(Select id from sysobjects where name='"&tablename&"') order by colid")
Set skiller=Conne.Execute("Select Type1, Type2 From Magic Where Magicnum="&Request.Querystring("skillno"))
If tablename="magic" Then
Set skillb=Conne.Execute("select * from magic where magicnum="&Request.Querystring("skillno")&"")
Else
Set skillb=Conne.Execute("select * from "&tablename&" where inum="&Request.Querystring("skillno"))
End If
col=0
Do While Not skillcolumn.Eof
Response.Write "<tr ><td class=""tdx"">"&skillcolumn(0)&":</td><td><input name="""&col&""" type=""text"" value="""&server.htmlencode(trim(skillb(col)))&""" class=""tvx""></td></tr>"&vbcrlf
col=col+1
Skillcolumn.MoveNext
Loop
%><tr><td colspan="2" align="right"><input type="submit" value="KAYDET" style="width:60px;font-size:10px"></td></tr></table>
</td><td valign="top" style="padding-left:50px;padding-top:50px;"><%Response.Write "<a href=""?nation="&nations&"&class="&classx&"&skilltur="&skilltur&"&skillno="&Request.Querystring("skillno")&"&table=magic""><blink>Magic</blink></a>"
if skiller("type1")<>0 Then
Response.Write "<br><br><a href=""?nation="&nations&"&class="&classx&"&skilltur="&skilltur&"&skillno="&Request.Querystring("skillno")&"&table=magic_type"&skiller("type1")&"""><blink>Magic Type"&skiller("type1")&"</blink></a>"
End If
if skiller("type2")<>0 Then
Response.Write "<br><br><a href=""?nation="&nations&"&class="&classx&"&skilltur="&skilltur&"&skillno="&Request.Querystring("skillno")&"&table=magic_type"&skiller("type2")&"""><blink>Magic Type"&skiller("type2")&"</blink></a>"
End If

End If

%></td>
</tr>
</table>