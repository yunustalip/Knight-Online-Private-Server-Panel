

<%
'*******************************************************
' Kodlarýmý kullandýðýnýz için teþekkürler
' Kullandýðýnýz siteyi bana bildirirseniz sevinirim
' Efkan 
' email :info@aywebhizmetleri.com
' web sayfalarýmý ziyaret etmeyi unutmayýnýz  
' http://www.makineteknik.com
' http://www.binbirkonu.com
' http://www.aywebhizmetleri.com
' http://www.tekrehberim.com
' http://www.hitlinkler.com
' Size uygun bir web sitem mutlaka vardýr ...
' LÜTFEN BU TÜR ÇALIÞMALARIN ÖNÜNÜ KESMEMEK ÝÇÝN TELÝF YAZILARINI SÝLMEYÝN
' EMEÐE SAYGI LÜTFEN 
' KÝÞÝSEL KULLANIM ÝÇÝN ÜCRETSÝZDÝR DÝÐER KULLANIMLARDA HAK TALEP EDÝLEBÝLÝR
'*******************************************************
%>





<div align="center">
<!--#INCLUDE file="forumayar.asp"-->

<% 
id= kontrol(temizle(request.querystring("id")))
pid= kontrol(temizle(request.querystring("pid")))
pid1= kontrol(temizle(request.querystring("pid1")))

sor = "SELECT  * FROM  grup WHERE altgrp="&id&" order by grp asc "
forum.Open sor,forumbag,1,3
If forum.eof Then
Response.Write "<BR><BR><B>Bu kayýd hiçbir kategoriye baðlý deðil</B>"
Response.End
End If
%>

<table width="100%" bgcolor="" bordercolor="<%=bgcolor2%>" border="3" cellspacing="3" cellpadding="0">
<tr>
<%i=1
do while not forum.eof  %>
<TD class="tdbrd1" bgcolor="<%=bgcolor1%>" align="center" valign="center" width="20%" height="20"onmouseover="this.style.backgroundColor='#CCFFFF';" onmouseout="this.style.backgroundColor='';">

<A HREF="default.asp?part=altgrp&id=<%=id%>&pid=<%=forum("id")%>">
<%=Left(forum("grp"),20)%>..</A>&nbsp;
</td>

<%if i mod 5 = 0 then
Response.Write "</tr><tr>"
End If
i = i + 1
forum.movenext  
loop 
forum.Close
%>
</tr></table>
</div>

<BR>
<div align="left">
<% sor = "SELECT  * FROM  grup WHERE id="&id&" "
forum.Open sor,forumbag,1,3%>
&nbsp;<B>
<A HREF="default.asp">Ana Sayfa</A>&nbsp;&gt;&gt;
<A HREF="default.asp?part=grp&id=<%=id%>"><%=forum("grp")%></B></A>

<%
grupadi = forum("grp")
forum.Close

if pid <> "" then
sor = "SELECT  * FROM  grup WHERE id="&pid&" "
forum.Open sor,forumbag,1,3%>
&gt;&gt;
<A HREF="default.asp?part=altgrp&id=<%=id%>&pid=<%=pid%>"><B><%=forum("grp")%></B></A>

<%
forum.Close
End If 

if pid1 <> "" And pid1<>0 then
sor = "SELECT  * FROM  grup1 WHERE id="&pid1&" "
forum3.Open sor,forumbag,1,3%>
&gt;&gt;
<A HREF="default.asp?part=altgrp&id=<%=id%>&pid=<%=pid%>&pid1=<%=forum3("id")%>">
<B><%=forum3("pidgrp")%></B></A>
<% forum3.close
else
End If

%>
</div>
