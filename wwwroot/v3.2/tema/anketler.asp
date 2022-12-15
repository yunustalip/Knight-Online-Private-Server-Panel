			<table border="0" width="550" id="table1" cellpadding="0" style="border-collapse: collapse">
				<tr>
					<td height="26" width="15">
					<img border="0" src="tema/images/blok_1.gif" width="15" height="26"></td>
					<td height="26" background="tema/images/blok_2.gif" width="525">
					<p align="center">
					<font class="baslik">Tüm Anketler</font></td>
					<td height="26" width="10">
					<img border="0" src="tema/images/blok_3.gif" width="10" height="26"></td>
				</tr>
				<tr>
					<td class="o_tab" valign="top" width="550" colspan="3" align="center">
<table border="0" width="530" id="table1" cellpadding="0" style="border-collapse: collapse">
	<tr>
		<td align="center"><font class="orta"><b>Eski Anketlere Oy Verilemez!</b></font></td>
	</tr>
	<tr>
		<td>
<%
set ab = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from ankets order by id desc"
ab.open SQL,data,1,3
Do While not ab.eof
%>
<table border="0" width="100%" id="table1" cellpadding="0" style="border-collapse: collapse">
	<tr>
		<td	align="center"><b><font class="orta"><%=ab("soru")%></font></b></td>
	</tr>
<!--Anket Þýklarý-->
<%
        Set tAnc = Server.CreateObject("ADODB.RecordSet") 
        tAnc.Open "anket Where a_id="&ab("id")&"",data,1,3
        Vote = 0
        Do While Not tAnc.EOF
        Vote = Vote + tAnc("deger")
        tAnc.MoveNext
        Loop
        tAnc.Close
        Set tAnc = NoThing

        Set tAnc = Server.CreateObject("ADODB.RecordSet")
        tAnc.Open "anket Where a_id="&ab("id")&"",data,1,3
Do While Not tAnc.EOF

                strOy = tAnc("deger")
                If strOy = "0" Then
                tOy = "0"
                Else
                tOy = (strOy /Vote) * 100
                End If
%>
		<tr>
			<td width="100%"><font class="blok"><%=tAnc("cevap")%> ( <%=strOy%> Oy - % <%=Left(tOy,4)%> )</font></td>
		</tr>
		<tr>
			<td width="100%">
<div style="width: 500px; height: 5px; border: 1px solid #000000">
<img src="tema/images/vote.gif" height="6" width="<%=Int(tOy)*5%>"></div>
			</td>
		</tr>
<% 
tAnc.MoveNext
Loop 
set toy = Server.CreateObject("ADODB.RecordSet")
SQL = "select SUM(deger) as oy_say from anket where a_id="&ab("id")&""
toy.open SQL,data,1,3
%>

		<tr>
			<td>
			<p align="center"><font class="blok">Toplam Oy: <%=toy("oy_say")%></font></p>
			</td>
		</tr>
<!--/Anket Þýklarý-->
<%
toy.close
set toy = Nothing
ab.movenext : loop
%>
</table>
</table>
					</td>
				</tr>
				<tr>
					<td colspan="3" width="550"><img src="tema/images/orta_alt.gif"></td>
				</tr>
			</table>