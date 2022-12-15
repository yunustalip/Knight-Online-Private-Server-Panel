<% Sub bulunamadi %>
			<table border="0" width="550" id="table1" cellpadding="0" style="border-collapse: collapse">
				<tr>
					<td height="26" width="15">
					<img border="0" src="tema/images/blok_1.gif" width="15" height="26"></td>
					<td height="26" background="tema/images/blok_2.gif" width="525">
					<p align="center">
					<font class="baslik">Kayýt Bulunamadý</font></td>
					<td height="26" width="10">
					<img border="0" src="tema/images/blok_3.gif" width="10" height="26"></td>
				</tr>
				<tr>
					<td class="o_tab" valign="top" width="550" colspan="3">
					<p align="center">
					<font class="orta">Kayýt Bulunamadý.<br>
					Yönetici Tarafýndan Silinmiþ Olabilir veya Yabancý Sayfalardan Gelmiþ Olabilirsiniz.
					<br>(Büyük Ýhtimal Kendin Salladýn Bu Sayfayý)</font></p>
					</td>
				</tr>
				<tr>
					<td colspan="3" width="550"><img src="tema/images/orta_alt.gif"></td>
				</tr>
			</table>
<% End Sub %>
<%
	Function ondalik(strWords) 
		strBadWords = Array(",") 
		strBadWordsReplace = Array("&#46;") 
			For iWords = 0 to uBound(strBadWords) 
				strWords = Replace(strWords, strBadWords(iWords), strBadWordsReplace(iWords),1,-1,1) 
			Next 
				If isNumeric(strWords) = True Then
					ondalik = int(strWords)
				Else
					ondalik = strWords 
				End if
	End Function

Sub Govde
X=Request.ServerVariables("SCRIPT_NAME")
if instr(x,"/blog.asp")>0 and Strseoayar="1" then
	Call Bulunamadi
Else
id=id1
	if id="" or isnumeric(id)=false then
		call bulunamadi
	else
set blogyo = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from blog where id="&id&""
blogyo.open SQL,data,1,3

if blogyo.eof then
call bulunamadi
else
if Request.Cookies("okunma")(id1) = "kapat" then
else
blogyo("hit")=blogyo("hit") + 1
blogyo.update
   Response.Cookies("okunma")(id) = "kapat"
   Response.Cookies("okunma").Expires = Now() + 1
end if
if blogyo("deger")="0" then
ortalama="0"
else
	ortalama = Round(blogyo("deger") / blogyo("degers"),1)
end if
	genislet = ortalama * 30
	ortalama = ondalik(ortalama)
	genislet = ondalik(genislet)
mesaj=Replace(blogyo("mesaj"),"{KES}","")
%>
			<table border="0" width="550" id="table1" cellpadding="0" style="border-collapse: collapse">
				<tr>
					<td height="26" width="15">
					<img border="0" src="tema/images/blok_1.gif" width="15" height="26"></td>
					<td height="26" background="tema/images/blok_2.gif" width="525">
					<p align="center">
					<font class="baslik">Blog</font></td>
					<td height="26" width="10">
					<img border="0" src="tema/images/blok_3.gif" width="10" height="26"></td>
				</tr>
				<tr>
					<td class="o_tab" valign="top" width="550" colspan="3">
<script type="text/javascript">
function ismaxlength(obj){
var mlength=obj.getAttribute? parseInt(obj.getAttribute("maxlength")) : ""
if (obj.getAttribute && obj.value.length>mlength)
obj.value=obj.value.substring(0,mlength)
}
</script>

<table border="0" width="530" height="24" id="table1" cellpadding="0" style="border-collapse: collapse" class="tool" align="center">
	<tr>
		<td width="172"></td>
		<td width="16">
		<img border="0" src="tema/images/rss-mini2.gif"></td>
		<td width="70"><a href="rss.asp?rss=blog&id=<%=id%>" target="_blank">RSS Takip</a></td>
		<td width="16">
		<img border="0" src="tema/images/oner.gif"></td>
		<td width="70"><a href="javascript:void(0)" ONCLICK="window.open('islem.asp?islem=tavsiyeet&id=<%=id%>','slide','top=20,left=20,width=250,height=250,toolbar=no,scrollbars=no');">Tavsiye Et</a></td>
		<td width="16">
		<img border="0" src="tema/images/yaz_word.gif"></td>
		<td width="70"><a target="_blank" href="islem.asp?islem=yazdir&id=<%=id%>&s=<%=session.sessionID%>">Ýndir (.doc)</a></td>
		<td width="10"></td>
		<td width="90"><font class="orta">Okunma: <%=blogyo("hit")%></font></td>
	</tr>
</table>

<div align="center">
<table border="0" width="530" id="table3" cellpadding="0" style="border-collapse: collapse" height="114">
	<tr>
		<td height="24">
		<font class="b_baslik"><%=blogyo("konu")%></font></td>
	</tr>
	<tr>
	   <td height="1" style="background-image: url('tema/images/nokta.gif'); background-repeat: repeat-x; background-position-y: bottom"></td>
	</tr>
	<tr>
		<td>
<br style="font-size:4px">
<center>
<table width="100%" id="table1" cellpadding="0" height="20" class="gecissf">
	<tr>
		<td align="left">
<%
Set Objonce = Server.CreateObject("adodb.RecordSet")
sql="SELECT id,konu From blog WHERE id < "&id&" order by id desc" 
Objonce.open sql, data, 1, 3
if not Objonce.eof then
%>
		<img border="0" src="tema/images/prev.gif"><a href="<%=SEOLink(objonce("id"))%>" title="Önceki Konu"> <%=Objonce("konu")%> </a>
<%
end if
objonce.close : set objonce = nothing 
%>
		</td>
		<td align="right">
<%
Set Objsonra = Server.CreateObject("adodb.RecordSet")
sql="SELECT id,konu From blog WHERE id > "&id&" order by id asc" 
Objsonra.open sql, data, 1, 3
if not Objsonra.eof then
%>
		<a href="<%=SEOLink(objsonra("id"))%>" title="Sonraki Konu"> <%=Objsonra("konu")%> </a><img border="0" src="tema/images/next.gif">
<%
end if
Objsonra.close : set Objsonra = nothing 
%>
		</td>
	</tr>
</table>
</center>
<br style="font-size:4px">
		</td>
	</tr>
	<tr>
		<td valign="top"><font class="orta"><%=mesaj%></font></td>
	</tr>
<%
Set etiketler = Server.CreateObject("adodb.RecordSet")
sql="select * From etiket where blog_id="&id&" order by id asc" 
etiketler.open sql, data, 1, 3
if not etiketler.eof then
%>
	<tr>
		<td><br><div class="etiket"><font class="orta">Etiketler: <%Do While Not Etiketler.eof%><a href="etiket.asp?etiket=<%=etiketler("etiket")%>"><%=etiketler("etiket")%></a>, &nbsp;<%etiketler.movenext : Loop%></font></div></td>
	</tr>
<%
end if
etiketler.close : set etiketler = nothing
%>
	<tr>
		<td>
		<!-- OYLAMA ALANI -->
			<div id="rating" 
				<%
					If Request.Cookies("Puan")(id) <> "" then
						response.write "style='display:none'"
					Else
						response.write "style='display:block'"
					End If
				%>>
				<ul class='star-rating'>
					<li class='current-rating' style='width:<%=genislet%>px;'></li>
					<li><a href="#" onclick="return xmlPost('islem.asp?islem=oyver&dokuman=<%=id1%>&oy=1&s=<%=session.sessionID%>','xmlspan')" title='5 üzerinden 1' class='one-star'>1</a></li>
					<li><a href="#" onclick="return xmlPost('islem.asp?islem=oyver&dokuman=<%=id1%>&oy=2&s=<%=session.sessionID%>','xmlspan')" title='5 üzerinden 2' class='two-stars'>2</a></li>
					<li><a href="#" onclick="return xmlPost('islem.asp?islem=oyver&dokuman=<%=id1%>&oy=3&s=<%=session.sessionID%>','xmlspan')" title='5 üzerinden 3' class='three-stars'>3</a></li>
					<li><a href="#" onclick="return xmlPost('islem.asp?islem=oyver&dokuman=<%=id1%>&oy=4&s=<%=session.sessionID%>','xmlspan')" title='5 üzerinden 4' class='four-stars'>4</a></li>
					<li><a href="#" onclick="return xmlPost('islem.asp?islem=oyver&dokuman=<%=id1%>&oy=5&s=<%=session.sessionID%>','xmlspan')" title='5 üzerinden 5' class='five-stars'>5</a></li>
				</ul>
			</div>
		<!-- OYLAMA BÝTTÝ -->
		<!-- OYLAMA SONUCU -->
		<div id="ratings"
			<%
				If Request.Cookies("Puan")(id) = "" then
					response.write "style='display:none'"
				Else
					response.write "style='display:block'"
				End If
			%>>
			<ul title='<%=blogyo("degers")%> oy aldý | Ortalama 5 üzerinden <%=Left(ortalama,3)%>'  class='star-rating'>
				<li class='current-rating' style='width:<%=genislet%>px;'></li>
				<li><strong class='one-star'></strong></li>
				<li><strong class='two-stars'></strong></li>
				<li><strong class='three-stars'></strong></li>
				<li><strong class='four-stars'></strong></li>
				<li><strong class='five-stars'></strong></li>
			</ul>
		</div>
		<span id="xmlspan"></span>

		</td>
	</tr>
<tr>
<td>
<a name="yorumlar"></a>
<div id="my_site_content">
</div>
<SCRIPT LANGUAGE=JAVASCRIPT>
function validate(form) {
if (form.Ekleyen.value == "") {
   alert("Adýnýzý ve Soyadýnýzý Yazýnýz.");
   return false; }
if (form.yorum.value == "") {
   alert("Lütfen Yorumunuzu Yazýnýz.");
   return false; }
return true;
}
</SCRIPT>

<SCRIPT language=JavaScript>
	function AddForm(form)
			{
				document.formcevap.yorum.value = document.formcevap.yorum.value + form
				document.formcevap.yorum.focus();
			}
</script>
<% if blogyo("yorumdurum")="0" then %>
<a name="yorum"></a>
<form method="post" action="islem.asp?islem=yorumekle&id=<%=id%>&s=<%=session.sessionID%>" onSubmit="return validate(this)" name="formcevap">
<div align="center">
<table border="0" id="table1" cellspacing="1" cellpadding="0" style="border-collapse: collapse" width="530">
	<tr>
		<td width="107">
		&nbsp;</td>
		<td width="440">
		<font class="orta"><b># Yorum Yaz #</b></font></td>
	</tr>
	<tr>
		<td width="107">
		<p align="right"><font class="orta">Ýsim :</font></td>
		<td width="440">
		<input type="text" class="alan" name="Ekleyen" size="26" value="<%=Request.Cookies("isim")%>"></td>
	</tr>
	<tr>
		<td width="107" valign="top">
		<p align="right"><font class="orta">Yorum :<%if session("admin")=false then%><br>(Max. 400 Karakter)<%End if%></font></td>
		<td width="440">
		<textarea rows="6" cols="39" class="alan" name="yorum"<%if session("admin")=false then%> onKeyUp="return ismaxlength(this)" maxlength="400"<%End if%>></textarea></td>
	</tr>
	<tr>
		<td width="110"></td>
		<td width="420">
<A href="javascript:AddForm(':)')"><img src="tema/images/smileys/smile.gif" border="0"></a> 
<A href="javascript:AddForm(':(')"><img src='tema/images/smileys/frown.gif' border=0></a> 
<A href="javascript:AddForm(':D')"><img src='tema/images/smileys/biggrin.gif' border=0></a> 
<A href="javascript:AddForm(':o:')"><img src='tema/images/smileys/redface.gif' border=0></a> 
<A href="javascript:AddForm(';)')"><img src='tema/images/smileys/wink.gif' border=0></a> 
<A href="javascript:AddForm(':p')"><img src='tema/images/smileys/tongue.gif' border=0></a> 
<A href="javascript:AddForm(':cool:')"><img src='tema/images/smileys/cool.gif' border=0></a> 
<A href="javascript:AddForm(':rolleyes:')"><img src='tema/images/smileys/rolleyes.gif' border=0></a> 
<A href="javascript:AddForm(':mad:')"><img src='tema/images/smileys/mad.gif' border=0></a> 
<A href="javascript:AddForm(':eek:')"><img src='tema/images/smileys/eek.gif' border=0></a> 
<A href="javascript:AddForm(':confused:')"><img src='tema/images/smileys/confused.gif' border=0>

</td>
	</tr>
	<tr>
		<td width="110">
		&nbsp;</td>
		<td width="420">
		<input type="submit" class="dugme" value="Gönder"></td>
	</tr>
</table>
</div>
</form>
<%
else
response.write "<center><font class=""orta""><b>Konu Yoruma Kapalý</b></font></center>"
end if
%>
</td>
</tr>
	<tr>
<td>
<%
benzer=blogyo("konu")
dim bnzr 
bnzr=""&benzer&"" 
bnzr = split(benzer," " )
konubenzer=Replace(bnzr(0),"'","")
Set benz = Server.CreateObject("adodb.RecordSet")
sql="SELECT id,konu,tarih,hit From blog WHERE Konu like '%"&konubenzer&"%' and id <> "&id&" order by konu asc" 
benz.open sql, data, 1, 3
if not benz.eof then
%>
<div align="center">
<table border="0" width="100%" cellpadding="0" style="border-collapse: collapse">
	<tr>
		<td class="tool" height="25" width="985" colspan="3"><font class="orta"><b>» Benzer 5 Konu</b></font></td>
	</tr>
	<tr>
		<td width="75%">
		<b><font class="orta">&nbsp;Konu Baþlýðý</font></b></td>
		<td width="15%">
		<b><font class="orta"> Tarih</font></b></td>
		<td width="10%">
		<b><font class="orta"> Okunma</font></b></td>
	</tr>
<% for z=1 to 5
if benz.eof then exit for %>
	<tr>
		<td width="75%"> <font class="orta">&nbsp;<a href="<%=SEOLink(benz("id"))%>"><%=benz("konu")%></a></font></td>
		<td width="15%"> <font class="orta"><%=benz("tarih")%></font></td>
		<td width="10%"> <font class="orta"><%=benz("hit")%></font></td>
	</tr>
<%
benz.MoveNext
next
%>
</table>
</div>
<%
end if
benz.close
set benz = nothing
%>
</td>
	</tr>
</table>
</div>
<%
blogyo.close
set blogyo=nothing
%>
					</td>
				</tr>
				<tr>
					<td colspan="3" width="550"><img src="tema/images/orta_alt.gif"></td>
				</tr>
			</table>
<%
End if
End if
End if
End Sub %>