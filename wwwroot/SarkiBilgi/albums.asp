<STYLE>
td{
color:#555555;
font-size: 12px;
font-family:verdana;
}
.sarkibilgi a{
color:#FFFFFF;
text-decoration:none;
}
.sarkibilgi a:hover{
	color: #ffffff;
	text-decoration:underline;
}

</STYLE><table><tr><td><a href="?sirala=sarkici">Þarkýcý</td><td><a href="?sirala=album">Albüm</a></td><td><a href="?sirala=sarki">Þarký</a></td>
<%
sirala=Request.Querystring("sirala")
if sirala="" Then
sirala="id"
End If
Set adocon = Server.CreateObject("ADODB.Connection")
adocon.open= "driver={SQL Server};server=localhost;uid=radyolarimadmin;pwd=864327142358;database=radyolarimdb" 

if Request.Querystring("islem")="sil" Then
id=Request.Querystring("id")
adocon.execute("delete musiclist where id="&id&"")
elseif Request.Querystring("islem")="resimsil" Then
id=Request.Querystring("id")
Set musiclist = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM musiclist where id="&id
musiclist.Open strSQL, adoCon ,1,3
musiclist("albumcover")=""
musiclist.update
End If

Function Base64Decode(ByVal base64String)
Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
Dim dataLength, sOut, groupBegin
 
base64String = Replace(base64String, vbCrLf, "")
base64String = Replace(base64String, vbTab, "")
base64String = Replace(base64String, " ", "")

dataLength = Len(base64String)
If dataLength Mod 4 <> 0 Then
Err.Raise 1, "Base64Decode", "Bad Base64 string."
Exit Function
End If

For groupBegin = 1 To dataLength Step 4
Dim numDataBytes, CharCounter, thisChar, thisData, nGroup, pOut
numDataBytes = 3
nGroup = 0

For CharCounter = 0 To 3

thisChar = Mid(base64String, groupBegin + CharCounter, 1)

If thisChar = "=" Then
numDataBytes = numDataBytes - 1
thisData = 0
Else
thisData = InStr(1, Base64, thisChar, vbBinaryCompare) - 1
End If
If thisData = -1 Then
Err.Raise 2, "Base64Decode", "Bad character In Base64 string."
Exit Function
End If

nGroup = 64 * nGroup + thisData
Next

nGroup = Hex(nGroup)

nGroup = String(6 - Len(nGroup), "0") & nGroup

pOut = Chr(CByte("&H" & Mid(nGroup, 1, 2))) + _
Chr(CByte("&H" & Mid(nGroup, 3, 2))) + _
Chr(CByte("&H" & Mid(nGroup, 5, 2)))

sOut = sOut & Left(pOut, numDataBytes)
Next

Base64Decode = sOut
End Function

Function StringToMultiByte(S)
Dim i, MultiByte
For i=1 To Len(S)
MultiByte = MultiByte & ChrB(Asc(Mid(S,i,1)))
Next
StringToMultiByte = MultiByte
End Function

Set musiclist =adocon.execute("SELECT * FROM musiclist ")


do while not musiclist.eof
if lenb(musiclist("albumcover"))>0 Then
%>
<tr>
<td><%=musiclist("sarkici")%></td>
<td><%=musiclist("album")%></td>
<td><%=musiclist("sarki")%></td>

<td><img src="x.asp?id=<%=musiclist("id")%>" width="70" height="70"></td>

<td><nobr><a href="?islem=sil&id=<%=musiclist("id")%>">Sil</a> <a href="?islem=resimsil&id=<%=musiclist("id")%>">Resim Sil</a></nobr></td></tr>
<%End If
musiclist.movenext
loop %>
</table>