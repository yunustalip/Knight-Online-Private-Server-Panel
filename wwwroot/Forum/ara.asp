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
if request("ara")="" then
Response.Write "<BR><BR><B>L�tfen aranacak kelimeyi belirtin...</B>"
Response.End
End If
ara=temizle(request("ara"))
nerde=temizle(request("nerde"))

If nerde="soru" then
sor = "select * from sorular WHERE onay=1 and baslik like '%"&ara&"%' OR onay=1 and aciklama like '%"&ara&"%' "
forum.Open sor,forumbag,1,3
Else
sor = "select * from cevaplar WHERE onay=1 and baslik like '%"&ara&"%' OR  onay=1 and aciklama like '%"&ara&"%' "
forum.Open sor,forumbag,1,3
End If

if forum.eof or forum.bof then
Response.Write "<BR><BR><BR><center><B>Kay�t bulunamad�... </B><P><a href=""javascript:history.back(1)""><B>&lt;&lt;Geri git</B></a>"
Response.End
End If
adet=forum.recordcount

shf = Request.QueryString("shf")
if shf="" then 
shf=1
End If
%>
<BR><BR>
<B><%=adet%> adet "<%=ara%>" bulundu..</B>


<!--D�K -->
<table width="100%" bgcolor="" bordercolor="" border="0" cellspacing="0" cellpadding="0">
<tr bgcolor="<%=bgcolor1%>" height="25"> 
<td align="left" width="85%" >&nbsp;&nbsp;<B>Ba�l�k</B></td>
<td align="center" width="15%" ><B>Tarih</B></td>
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
&nbsp;<IMG SRC="images/blank.gif" WIDTH="9" HEIGHT="7" BORDER=0 ALT="">

<%If nerde="soru" then%>
<A HREF="?part=oku&id=<%=forum("grp")%>&pid=<%=forum("altgrp")%>&urun=<%=forum("id")%>">
<%=buyukharf(forum("baslik"))%></a>
<%else%>
<A HREF="?part=oku&id=<%=forum("grp")%>&pid=<%=forum("altgrp")%>&urun=<%=forum("soruid")%>">
<%=buyukharf(forum("baslik"))%></a>
<%End If%>
</td>

<TD align="center" valign="center"><%=forum("tarih")%></td>
</tr>
<% renk=renk + 1
forum.movenext
Next
forum.close %>
</table>
<!--D�K SON -->



<P>
Sayfalar :
<%
say=0
for y=1 to sayfa 
if say mod 10 = 0 then
Response.Write "<BR>"
End If
if  y=cint(shf ) then 
Response.Write "<B>["&y&"]</B>"
else
Response.Write "<a href='default.asp?part=ara&ara="&ara&"&nerde="&nerde&"&shf="&y&"'>["&y&"]</a>"
End If
say=say+1
next
set forum =nothing

%>
<P>




