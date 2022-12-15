<!--#INCLUDE file="forumayar.asp"-->
<%
If Session("efkanlogin")=True <> True Then 
Response.Write "<script language='JavaScript'>alert('Bu alana girmeye yetkiniz yoktur...');</script>"
Response.Write "<meta http-equiv='Refresh' content='0; URL=default.asp?part=uyegorev'>"
Response.End
End If

gel = Request.ServerVariables("SCRIPT_NAME")
gorev = temizle(request.querystring("gorev"))

id= kontrol(temizle(request.querystring("id")))
pid= kontrol(temizle(request.querystring("pid")))
urun= kontrol(temizle(request.querystring("urun")))
cevapid=kontrol(request.querystring("cevapid"))



If gorev="ac" Then 
sor = "select * from sorular WHERE id="&urun&""
forum.Open sor,forumbag,1,3
forum("acik")=1
forum.update
forum.close
Response.Redirect	"?part=oku&id="&id&"&pid="&pid&"&urun="&urun&""
End If

If gorev="kapa" Then 
sor = "select * from sorular WHERE id="&urun&""
forum.Open sor,forumbag,1,3
forum("acik")=0
forum.update
forum.close
Response.Redirect	"?part=oku&id="&id&"&pid="&pid&"&urun="&urun&""
End If


If gorev="sorusil" Then 
sor = "DELETE from sorular WHERE id="&urun&""
forum.Open sor,forumbag,1,3
sor = "DELETE from cevaplar WHERE soruid="&urun&""
forum1.Open sor,forumbag,1,3
Response.Redirect	"?part=altgrp&id="&id&"&pid="&pid&""
End If


If gorev="cevapsil" Then 
sor = "DELETE from cevaplar WHERE id="&cevapid&""
forum.Open sor,forumbag,1,3
Response.Redirect	"?part=oku&id="&id&"&pid="&pid&"&urun="&urun&""
End If


if gorev="tasi" Then 
id =Trim(request.querystring("id"))
anakat = request("anakat")
altkat = request("altkat")
if  anakat="" Then %>
<BR><BR>
<B>Seçilen Kaydý Hangi Ana kategoriye taþýyacaksýnýz...</B><P>
<table width="50%" bgcolor="" bordercolor="" border="0" cellspacing="0" cellpadding="1">
<tr height="60"><td align="center">
<form method="POST" action="?part=gorev&gorev=tasi&id=<%=id%>">
<B>Kayýdýn olmasýný istediðiniz Ana Kategoriyi seçiniz :</B><P>
<select NAME="anakat">
<option value="" selected>Lütfen Ana Kategori seçiniz</option>
<%sor = "Select * from grup where altgrp =0 order by grp asc" 
forum.Open sor,forumbag,1,3
do while not forum.eof  %>
<option value="<%=forum("id")%>"><%=forum("grp")%></option>
<%forum.movenext 
loop 
forum.Close%>
</select>
<input type="submit" value="Tamam" name="submit" >
</form></td></tr></table>
<% ElseIf anakat<>"" And altkat="" Then %><BR><BR>
<B>Seçilen Kaydý Hangi Alt  kategoriye taþýyacaksýnýz...</B><P>
<table width="50%" bgcolor="" bordercolor="" border="0" cellspacing="0" cellpadding="1">
<tr height="60"><td align="center">
<form method="POST" action="?part=gorev&gorev=tasi&id=<%=id%>&anakat=<%=anakat%>">
<B> Kayýdýn olmasýný istediðiniz Alt Kategoriyi seçiniz :</B><P>
<select NAME="altkat">
<option value="" selected>Lütfen Alt Kategori seçiniz</option>
<%sor = "Select * from grup where altgrp ="& anakat &"  order by grp asc" 
forum.Open sor,forumbag,1,3
do while not forum.eof  %>
<option value="<%=forum("id")%>"><%=forum("grp")%></option>
<%forum.movenext 
loop 
forum.Close%>
</select>
<input type="submit" value="Tamam" name="submit" >
</form></td></tr></table>
<%Else
sor  = "select * from sorular where id="&id&" "
forum.Open sor,forumbag,1,3
forum("grp") = anakat
forum("altgrp") = altkat
forum.Update
forum.close
sor  = "select * from cevaplar where soruid ="&id&" "
forum.Open sor,forumbag,1,3
If forum.eof Then
else
forum("grp") = anakat
forum("altgrp") = altkat
forum.Update
End If
'Response.Write "<script language='JavaScript'>alert('Kayýt ve Bu kayýta ait mesajlar taþýndý');</script>"
'Response.Write "<meta http-equiv='Refresh' content='1; URL=?part=grp'>" 
Response.Redirect	"?part=oku&id="&forum("grp")&"&pid="&forum("altgrp")&"&urun="&forum("soruid")&""
forum.close
End If
End If




if gorev="subtasi" Then 
urun =Trim(request.querystring("urun"))
pid =Trim(request("pid"))
pid1 =Trim(request.form("pid1"))

If pid1="" Then %>
<form method="POST" action="?part=gorev&gorev=subtasi&urun=<%=urun%>">
<BR><BR><BR>
<B>Kayýdýn olmasýný istediðiniz Sub Kategoriyi seçiniz :</B><P>
<select NAME="pid1">
<option value="" selected>Lütfen Sub Kategori seçiniz</option>
<%sor = "Select * from grup1 where pid="&pid&" order  by pidgrp asc"  
forum1.Open sor,forumbag,1,3
do while not forum1.eof %>
<option value="<%=forum1("id")%>"><%=forum1("pidgrp")%></option>
<%forum1.movenext 
loop 
forum1.Close%>
</select>
<input type="submit" value="Tamam" name="submit" >
</form>
<% Else 

sor  = "select * from sorular where id ="&urun&" "
forum.Open sor,forumbag,1,3
forum("sub")=Trim(request.form("pid1"))
forum.update
Response.Redirect	"?part=oku&id="&forum("grp")&"&pid="&forum("altgrp")&"&pid1="&forum("sub")&"&urun="&forum("id")&""
forum.close
End If
End If
 



%>