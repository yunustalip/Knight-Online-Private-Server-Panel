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

<!--#INCLUDE file="forumayar.asp"-->
<div align="center">
<% 
Response.Buffer = True

If  Session("efkanlogin")<>True  Then
Response.Write "Bu i�lem i�in yetkiniz yok"
Response.End
End If 


neresi=request.querystring("neresi")
gorev=request.querystring("gorev")


If gorev="onayli" Then
    If neresi="sorular" then
    sor = "select * from sorular WHERE onay=1 order by id desc "
	ElseIf neresi="cevaplar" then
    sor = "select * from cevaplar WHERE onay=1 order by id desc "
	End If
    forum.Open sor,forumbag,1,3
    adet=forum.recordcount

if forum.eof  then
Response.Write "<BR><BR><BR><center><B>Onay Bekleyen Yok... </B><P><a href=""javascript:history.back(1)""><B>&lt;&lt;Geri git</B></a>"
Response.End
End If


shf = Request.QueryString("shf")
if shf="" then 
shf=1
End If %>


<B><%=gorev%> &nbsp;<%=neresi%></B>
<table width="100%" bgcolor="" bordercolor="" border="0" cellspacing="0" cellpadding="0">
<tr bgcolor="<%=bgcolor1%>" height="25"> 
<td align="left" width="60%" >&nbsp;&nbsp;<B>Ba�l�k</B></td>
<td align="center" width="10%" ><B>Tarih</B></td>
<td align="center" width="20%" ><B>��lem</B></td>
</tr>
<% renk = 0
forum.pagesize =40
forum.absolutepage = shf
sayfa = forum.pagecount
for i=1 to forum.pagesize
if forum.eof then exit for
if renk mod 2 then
bgcolor = bgcolor1
else
bgcolor = bgcolor2
End If
%>
<tr bgcolor="<%=bgcolor%>"  height="20">
<TD align="left" valign="center">
&nbsp;<IMG SRC="images/blank.gif" WIDTH="9" HEIGHT="7" BORDER="0" ALT="">

<A HREF="?part=onayli&gorev=oku&id=<%=forum("id")%>&neresi=<%=neresi%>">
<%=buyukharf(forum("baslik"))%></a>
</td>

<TD align="center" valign="center"><%=forum("tarih")%></td>

<TD align="center" valign="center">

<A HREF="?part=onayli&gorev=oku&id=<%=forum("id")%>&neresi=<%=neresi%>">
<IMG SRC="forumimg/oku.gif" WIDTH="31" HEIGHT="17" BORDER="0" ALT=""></a>

<%If request.querystring("neresi")="sorular" Then%>
<A HREF="?part=onayli&gorev=sorusil&id=<%=forum("id")%>&neresi=<%=neresi%>">
<IMG SRC="forumimg/ksil.gif" WIDTH="55" HEIGHT="17" BORDER="0" ALT="Konuyu ve buna ait mesajlar� siler"></a>
<%elseIf request.querystring("neresi")="cevaplar" then%>
<A HREF="?part=onayli&gorev=cevapsil&id=<%=forum("id")%>&neresi=<%=neresi%>">
<IMG SRC="forumimg/msil.gif" WIDTH="55" HEIGHT="17" BORDER="0" ALT="Sadece mesaj� siler"></a>
<%End If%>

</td>

</tr>
<% renk=renk + 1
forum.movenext
Next
forum.close %>
</table>
<!--D�K SON -->

<P>Sayfalar :
<%
say=0
for y=1 to sayfa 
if say mod 10 = 0 then
Response.Write "<BR>"
End If
if  y=cint(shf ) then 
Response.Write "<B>["&y&"]</B>"
else
Response.Write "<a href='default.asp?part=onayli&gorev="&gorev&"&neresi="&neresi&"&shf="&y&"'>["&y&"]</a>"
End If
say=say+1
Next

End If




If gorev="oku" Then
id=Trim(request.querystring("id"))
neresi=Trim(request.querystring("neresi"))
   If neresi="sorular" then
   sor = "select  * from  sorular where id="&id&"  "
   elseIf neresi="cevaplar" then
   sor = "select  * from  cevaplar where id="&id&"  "
    End If
    forum.Open sor,forumbag,1,3

%>

<table width="100%" border="1" bgcolor="" bordercolor="#FFFFFF" align="center" cellpadding="0" cellspacing="0">
<tr height="">
<td width="100%" align="left" valign="center">
<CENTER>


<%If neresi="sorular" Then%>
<A HREF="?part=onayli&gorev=sorusil&id=<%=forum("id")%>&neresi=<%=neresi%>">
<IMG SRC="forumimg/ksil.gif" WIDTH="55" HEIGHT="17" BORDER="0" ALT="Konuyu ve buna ait mesajlar� siler"></a>
<%else%>
<A HREF="?part=onayli&gorev=cevapsil&id=<%=forum("id")%>&neresi=<%=neresi%>">
<IMG SRC="forumimg/msil.gif" WIDTH="55" HEIGHT="17" BORDER="0" ALT="Sadece mesaj� siler"></a>
<%End If%>


<P>

<B><%=forum("baslik")%></B></CENTER>
<P>
<%=forum("aciklama")%>
<P>


<B>Kim yazd� :</B>
<A HREF="default.asp?part=uyegorev&gorev=uyebilgi&id=<%=forum("uyeid")%>"><%=forum("kadi")%></a>
<BR>
<B>�ye No :</B><%=forum("uyeid")%><BR>
<B>�pno :</B><%=forum("ipno")%><BR>
</td></tr></table>

<%
End If



If gorev="sorusil" Then 
id=request.querystring("id")

sor = "DELETE from sorular WHERE id="&id&""
forum.Open sor,forumbag,1,3
sor = "DELETE from cevaplar WHERE soruid="&id&""
forum1.Open sor,forumbag,1,3
Response.Redirect	"?part=onayli&gorev=onayli&neresi="&neresi&""
End If


If gorev="cevapsil" Then 
id=request.querystring("id")

sor = "DELETE from cevaplar WHERE id="&id&""
forum.Open sor,forumbag,1,3
Response.Redirect	"?part=onayli&gorev=onayli&neresi="&neresi&""
End If



set forum =nothing

%>
<P>



