<%@CODEPAGE="1254" LANGUAGE="VbScript" LCID="1055"%>

<%
'*******************************************************
' Kodlar�m� kulland���n�z i�in te�ekk�rler
' Kulland���n�z siteyi bana bildirirseniz sevinirim
' Efkan 
' email :info@aywebhizmetleri.com
' web sayfalar�m� ziyaret etmeyi unutmay�n�z  
' http://www.makineteknik.com
' http://www.binbirkonu.com
' http://www.aywebhizmetleri.com
' http://www.tekrehberim.com
' http://www.hitlinkler.com
' Size uygun bir web sitem mutlaka vard�r ...
' L�TFEN BU T�R �ALI�MALARIN �N�N� KESMEMEK ���N TEL�F YAZILARINI S�LMEY�N
' EME�E SAYGI L�TFEN 
' K���SEL KULLANIM ���N �CRETS�ZD�R D��ER KULLANIMLARDA HAK TALEP ED�LEB�L�R
'*******************************************************
%>

<%








%>


<!--#INCLUDE file="forumayar.asp"-->


<div align="center">

<BR><B>SAYAC �STAT�ST�KLER�</B><P>

<A HREF="default.asp?part=sayac&sayac=gunhit"><B>G�nl�k Hit D�k�m�</B></A> |
<A HREF="default.asp?part=sayac&sayac=saysite"><B>Hit G�nderen Siteler</B></A> |
<A HREF="default.asp?part=sayac&sayac=gunluk"><B>G�nl�k Gelen D�k�m�</B></A> |
<A HREF="default.asp?part=sayac&sayac=ip"><B>�p D�k�m�</B></A> |
<A HREF="default.asp?part=sayac&sayac=yasakli"><B>Yasakl�lar</B></A> |
<A HREF="default.asp?part=sayac&sayac=yasakliekle"><B>Yasaklanacak �p Ekle</B></A> 
<P>


<form name="temizle">
<select name="menu">
<option value="default.asp?part=sayac&sayac=iptemizle">5 g�n �nceki ipleri sil</option>
<option value="default.asp?part=sayac&sayac=gunluktemizle">5 g�n �nceki G�nl�k Siteleri Sil</option>
<option value="default.asp?part=sayac&sayac=saysitetemizle">1 ay �nceki hit g�nderenleri Sil</option>
</select>
<input type="button" onClick="location=document.temizle.menu.options[document.temizle.menu.selectedIndex].value;" value=" Temizle ">
</form>


<% 
Response.Buffer = True 

sayac=request.querystring("sayac")

If sayac="" Or sayac="gunhit" Then 
sor="select * from say_hit order by id  desc   " 
efkan.Open sor,Sur,1,3
adet=efkan.recordcount

if efkan.eof or efkan.bof then
bilgiver1("Kay�t Bulunamad�.")
else

shf = Request.QueryString("shf")
if shf="" then 
shf=1
End If %>

<B>G�N H�TLER�</B>
<table background="" width="65%" bgcolor="" bordercolor="#FFFFFF" border="0" cellspacing="0" cellpadding="2">
<tr bgcolor="<%=bgcolor1%>">
<td align="center"  width="30%"><B>Tarih</B></td>
<td align="center" width="30%"><B>Tekil</B></td>
<td align="center"  width="30%"><B>�o�ul</B></td>
</tr>
<%
efkan.pagesize =50
efkan.absolutepage = shf
sayfa = efkan.pagecount
for i=1 to efkan.pagesize
if efkan.eof then exit For
if renk mod 2 then
bgcolor = bgcolor1
else
bgcolor = bgcolor2
End If 
%>
<tr bgcolor="<%=bgcolor%>" onmouseover="this.style.backgroundColor='#CCFFFF';" onmouseout="this.style.backgroundColor='';">
<td align="center"><%= efkan("gun")%></td>
<td align="center"><%= efkan("tekil")%></td>
<td align="center"><%= efkan("cogul")%></td>
</tr>
<%
renk=renk + 1
efkan.movenext
next 
efkan.close
%>
</table>
<P>
Sayfalar :
<%say=0
for y=1 to sayfa 
if say mod 20 = 0 then
Response.Write "<BR>"
End If
if  y=cint(shf) then 
Response.Write "<B>["&y&"]</B>"
else
Response.Write "<a href='default.asp?part=sayac&sayac="&sayac&"&shf="&y&"'>["&y&"]</a>"
End If
say=say+1
Next
End If
End If


'/////////////////////// H�T G�NDERENLER�N TOPLAM SAYACI /////////////////////////
If sayac="saysite"  Then 
sor="select * from say_site order by hit desc   " 
efkan.Open sor,Sur,1,3
adet=efkan.recordcount

if efkan.eof or efkan.bof then
bilgiver1("Kay�t Bulunamad�.")
else
shf = Request.QueryString("shf")
if shf="" then 
shf=1
End If
%>

<B>H�T G�NDEREN S�TELER</B>
<table background="" width="65%" bgcolor="" bordercolor="#FFFFFF" border="0" cellspacing="0" cellpadding="2">
<tr bgcolor="<%=bgcolor1%>">
<td align="center"  width="20%"><B>En son</B></td>
<td align="center" width="20%"><B>Hit</B></td>
<td align="center"  width="60%"><B>Hit G�nderen</B></td>
</tr>
<%
efkan.pagesize =50
efkan.absolutepage = shf
sayfa = efkan.pagecount
for i=1 to efkan.pagesize
if efkan.eof then exit for
if renk mod 2 then
bgcolor = bgcolor1
else
bgcolor = bgcolor2
End If 
%>
<tr bgcolor="<%=bgcolor%>" onmouseover="this.style.backgroundColor='#CCFFFF';" onmouseout="this.style.backgroundColor='';">
<td align="center"><%= efkan("gun")%></td>
<td align="center"><%= efkan("hit")%></td>
<td align="left"><A HREF="<%= efkan("site_name")%>" target="_blank"><%= efkan("site_name")%></A></td>
</tr>

<%
renk=renk+1
efkan.movenext
next 
efkan.close%>
</table>
Sayfalar :
<%say=0
for y=1 to sayfa 
if say mod 20 = 0 then
Response.Write "<BR>"
End If
if  y=cint(shf) then 
Response.Write "<B>["&y&"]</B>"
else
Response.Write "<a href='default.asp?part=sayac&sayac="&sayac&"&shf="&y&"'>["&y&"]</a>"
End If
say=say+1
Next
End If
End If


'/////////////////////// G�NL�K GELENLER�N D�K�M� /////////////////////////
If sayac="gunluk"  Then 
sor="select * from site_gel order by id desc   " 
efkan.Open sor,Sur,1,3
adet=efkan.recordcount
if efkan.eof or efkan.bof then
bilgiver1("Kay�t Bulunamad�.")
else

shf = Request.QueryString("shf")
if shf="" then 
shf=1
End If
%>
<B>G�NL�K H�T G�NDERENLER�N D�K�M�</B>
<table background="" width="65%" bgcolor="" bordercolor="#FFFFFF" border="0" cellspacing="0" cellpadding="2">
<tr bgcolor="<%=bgcolor1%>">
<td align="center"  width="20%" ><B>G�n</B></td>
<td align="center" width="20%" ><B>Hit</B></td>
<td align="center"  width="60%" ><B>Url</B></td>
</tr>
<%
efkan.pagesize =50 
efkan.absolutepage = shf
sayfa = efkan.pagecount
for i=1 to efkan.pagesize
if efkan.eof then exit For
if renk mod 2 then
bgcolor = bgcolor1
else
bgcolor = bgcolor2
End If 
%>
<tr bgcolor="<%=bgcolor%>" onmouseover="this.style.backgroundColor='#CCFFFF';" onmouseout="this.style.backgroundColor='';">
<td align="center"><%= efkan("gun")%></td>
<td align="center"><%= efkan("hit")%></td>
<td align="left">
<A HREF="<%= efkan("site_gel")%>" target="_blank"><%= efkan("site_gel")%></A>
</td>
</tr>
<%
renk=renk+1
efkan.movenext
next 
efkan.close
%>
</table>

Sayfalar :
<%say=0
for y=1 to sayfa 
if say mod 20 = 0 then
Response.Write "<BR>"
End If
if  y=cint(shf) then 
Response.Write "<B>["&y&"]</B>"
else
Response.Write "<a href='default.asp?part=sayac&sayac="&sayac&"&shf="&y&"'>["&y&"]</a>"
End If
say=say+1
Next
End If
End If


'/////////////////////// �P D�K�M� /////////////////////////
If sayac="ip"  Then 
sor="select * from say_ip order by id desc   " 
efkan.Open sor,Sur,1,3
adet=efkan.recordcount
if efkan.eof or efkan.bof then
bilgiver1("Kay�t Bulunamad�.")
else

shf = Request.QueryString("shf")
if shf="" then 
shf=1
End If
%>
<B>Z�YARET EDEN �P NUMARALARI</B>
<table background="" width="65%" bgcolor="" bordercolor="#FFFFFF" border="0" cellspacing="0" cellpadding="2">
<tr bgcolor="<%=bgcolor1%>">
<td align="center"  width="20%" ><B>En Son</B></td>
<td align="center" width="20%" ><B>Hit</B></td>
<td align="center"  width="60%" ><B>�p No</B></td>
</tr>
<%
efkan.pagesize =50 
efkan.absolutepage = shf
sayfa = efkan.pagecount
for i=1 to efkan.pagesize
if efkan.eof then exit For
if renk mod 2 then
bgcolor = bgcolor1
else
bgcolor = bgcolor2
End If 
%>
<tr bgcolor="<%=bgcolor%>" onmouseover="this.style.backgroundColor='#CCFFFF';" onmouseout="this.style.backgroundColor='';">
<td align="center"><%= efkan("vakit")%></td>
<td align="center"><%= efkan("hit")%></td>
<td align="left"><%= efkan("ip_number")%></td>
</tr>
<%
renk=renk+1
efkan.movenext
next 
efkan.close
%>
</table>
Sayfalar :
<%say=0
for y=1 to sayfa 
if say mod 20 = 0 then
Response.Write "<BR>"
End If
if  y=cint(shf) then 
Response.Write "<B>["&y&"]</B>"
else
Response.Write "<a href='default.asp?part=sayac&sayac="&sayac&"&shf="&y&"'>["&y&"]</a>"
End If
say=say+1
Next
End If
End If


If sayac="iptemizle"  Then 

sor="SELECT * FROM say_ip  "
efkan.Open sor,Sur,1,3
do while not efkan.eof  
zaman=datediff("d",efkan("vakit"),now)  
if zaman > 5 then
sor="DELETE from say_ip WHERE id = "&efkan("id")&"  "
efkan1.Open sor,Sur,1,3
End If
efkan.movenext
Loop
efkan.Close
Response.Redirect "default.asp?part=sayac&sayac=ip"
End If


If sayac="gunluktemizle"  Then 

sor="SELECT * FROM site_gel  "
efkan.Open sor,Sur,1,3
do while not efkan.eof  
zaman=datediff("d",efkan("gun"),now)  ' 1 ay �ncesi
if zaman > 5 then
sor="DELETE from site_gel WHERE id = "&efkan("id")&"  "
efkan1.Open sor,Sur,1,3
End If
efkan.movenext
Loop
efkan.Close
Response.Redirect "default.asp?part=sayac&sayac=gunluk"
End If


If sayac="saysitetemizle"  Then 

sor="SELECT * FROM say_site  "
efkan.Open sor,Sur,1,3
do while not efkan.eof  
zaman=datediff("m",efkan("gun"),now)  ' 1 ay �ncesi
if zaman > 1 then
sor="DELETE from say_site WHERE id = "&efkan("id")&"  "
efkan1.Open sor,Sur,1,3
End If
efkan.movenext
Loop
efkan.Close
Response.Redirect "default.asp?part=sayac&sayac=saysite"
End If





'//////// YASAKLILAR
if sayac="yasakli" then  %>
<A HREF="?sayac=yasakliekle"><B>Yasaklanacak �p Ekle</B></A><P>
<% sor = "Select * from yasakli order by id desc " 
efkan.Open sor,Sur,1,3
adet=efkan.recordcount
if efkan.eof or efkan.bof then
Response.Write "Yasakl� Yok"
Else
shf = Request.QueryString("shf")
if shf="" then 
shf=1
End If %>
<B>YASAKLI �P �PLER</B><BR>
<table width="50%" bgcolor="#F9F9F9" bordercolor="#FFFFFF" border="1" cellspacing="0" cellpadding="3">
<tr bgcolor="<%=bgcolor%>">
<td width="1%"><B>id</B></td>
<td width="10%" align="center"><B>Tarih</B></td>
<td width="10%" align="center"><B>�p</B></td>
<td width="4%" align="center"><B>��lem</B></td>
</tr>
<% efkan.pagesize =50
efkan.absolutepage = shf
sayfa = efkan.pagecount
for i=1 to efkan.pagesize
if efkan.eof then exit For
if renk mod 2 then
bgcolor = bgcolor1
else
bgcolor = bgcolor2
End If %>
<tr bgcolor="<%=bgcolor%>" onmouseover="this.style.backgroundColor='#CCFFFF';" onmouseout="this.style.backgroundColor='';">
<td align="center"><%=efkan("id")%></td>
<td align="center"><%=efkan("tarih")%></td>
<td align="center"><%=efkan("ip")%></td>
<td align="center">
<A HREF="default.asp?part=sayac&sayac=yasaklisil&id=<%=efkan("id")%>">Sil</A>
</td></tr>
<%
renk=renk+1
efkan.movenext
Next
efkan.close%>
</table><P>
Sayfalar :
<%say=0
for y=1 to sayfa 
if say mod 20 = 0 then
Response.Write "<BR>"
End If
if  y=cint(shf) then 
Response.Write "<B>["&y&"]</B>"
Else
	  Response.Write "<a href='?sayac=yasakli&shf="&y&"'>["&y&"]</a>"
End If
say=say+1
next
End If
End If




'////////////////////////// YASAKLI EKLE /////////////////////////////////
If sayac="yasakliekle" Then %>
<form method="POST" action="default.asp?part=sayac&sayac=yasakliekle">
<A HREF="?sayac=yasakli">T�m Yasakl�lar</A>
<table width="50%" bgcolor="" bordercolor="#f5f5ff" border="1" cellspacing="0" cellpadding="3">
<tr><td width="40%">Yasaklanacak �p</td><td width="60%">
<input name="ip" type="text" value="" size="40"></td></tr>
<tr><td align="center" colspan="2">
<input type="submit" value=" Ekle " name="submit" > <INPUT TYPE="reset" value=" Temizle ">
</td></tr></table></form>
<% 
if request.form("ip")=""  then
else
sor = "Select * from yasakli  " 
efkan.Open sor,Sur,1,3
efkan.AddNew
  efkan("ip")         =Trim(request.form("ip"))
  efkan("tarih")     =Now()
efkan.Update
efkan.close
Response.Redirect "default.asp?part=sayac&sayac=yasakli"
End If
End If






'//////////////////YASAKLI S�L
if sayac="yasaklisil" then 

id=request.querystring("id")
sor="DELETE from yasakli WHERE id = "&id&"  "
efkan.Open sor,sur,1,3
Response.Redirect "default.asp?part=sayac&sayac=yasakli"
End If




Set efkan1=Nothing
Set efkan=Nothing
%>


<P>





