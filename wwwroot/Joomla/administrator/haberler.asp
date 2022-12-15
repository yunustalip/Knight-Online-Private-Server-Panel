<%
'      JoomlASP Site Yönetimi Sistemi (CMS)
'
'      Copyright (C) 2007 Hasan Emre ASKER
'
'      This program is free software; you can redistribute it and/or modify it
'      under the terms of the GNU General Public License as published by the Free
'      Software Foundation; either version 3 of the License, or (at your option)
'      any later version.
'
'      This program is distributed in the hope that it will be useful, but WITHOUT
'      ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
'      FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
'
'      You should have received a copy of the GNU General Public License along with
'      this library; if not, write to the JoomlASP Asp Yazýlým Sistemleri., Kargaz Doðal Gaz Bilgi Ýþlem Müdürlüðü
'       36100 Kars / Merkez 
'		Tel : 0544 275 9804 - 0537 275 3655
'		Mail: emre06@hotmail.com.tr / info@joomlasp.com/.net/.org
'
'
'		Lisans Anlaþmasý Gereði Lütfen Google Reklam Bölümünü Sitenizden kaldýrmayýnýz. Bu sizin GOOGLE reklamlarýný yapmanýza
'		kesinlikle bir engel deðildir. reklam.asp içeriðinin yada yayýnladýðý verinin deðiþmesi lisans politikasýnýn dýþýna çýkýlmasýna
'		ve JoomlASP CMS sistemini ücretsiz yayýnlamak yerine ücretlie hale getirmeye bizi teþfik etmektedir. Bu Sistem için verilen emeðe
'		saygý ve bir çeþit ödeme seçeneði olarak GOOGLE reklamýmýzýn deðiþtirmemesi yada silinmemesi gerekmektedir.
%>
<td height="25" colspan="3" bgcolor="#333333"><span class="style2">&nbsp;JoomlASP'den Haberler </span></td>
        </tr>
      <tr>
        <td height="20" bgcolor="#666666"><span class="style3"> &nbsp;Yeni Modüller </span></td>
        <td height="20" bgcolor="#666666"><span class="style3"> &nbsp;Yeni Bileþenler </span></td>
        <td height="20" bgcolor="#666666"><span class="style3"> &nbsp;Versiyon Kontrol </span></td>
      </tr>
      <tr>
        <td valign="top" bgcolor="#999999"><%
URL = "http://www.joomlasp.com/haberler/rss.asp?istek=modul"

set xmlDoc = createObject("MSXML.DOMDocument")
xmlDoc.async = false
xmlDoc.setProperty "ServerHTTPRequest", true
xmlDoc.load(URL)

If (xmlDoc.parseError.errorCode <> 0) then
    Response.Write "XML Hatasý: " & xmlDoc.parseError.reason
Else

    set channelNodes = xmlDoc.selectNodes("//item/*")
    for each entry in channelNodes
        if entry.tagName = "title" then
                strBaslik = entry.text
		elseif entry.tagName = "link" then
		strLink = entry.text 
		elseif entry.tagName = "description" then
		strBilgi = entry.text        
	Response.Write "&nbsp;<a target=""_blank"" href="&strLink&">"&strBaslik&"</a><br>"
	end if
    next
end If

%></td>
        <td valign="top" bgcolor="#999999"><%
URL = "http://www.joomlasp.com/haberler/rss.asp?istek=bilesen"

set xmlDoc = createObject("MSXML.DOMDocument")
xmlDoc.async = false
xmlDoc.setProperty "ServerHTTPRequest", true
xmlDoc.load(URL)

If (xmlDoc.parseError.errorCode <> 0) then
    Response.Write "XML Hatasý: " & xmlDoc.parseError.reason
Else

    set channelNodes = xmlDoc.selectNodes("//item/*")
    for each entry in channelNodes
        if entry.tagName = "title" then
                strBaslik = entry.text
		elseif entry.tagName = "link" then
		strLink = entry.text 
		elseif entry.tagName = "description" then
		strBilgi = entry.text        
	Response.Write "&nbsp;<a target=""_blank"" href="&strLink&">"&strBaslik&"</a><br>"
	end if
    next
end If

%></td>
        <td valign="top" bgcolor="#999999">&nbsp;<b>Mevcut Sürüm:</b> <%=surum%><br />
          <%
URL = "http://www.joomlasp.com/haberler/rss.asp?istek=versiyon"

set xmlDoc = createObject("MSXML.DOMDocument")
xmlDoc.async = false
xmlDoc.setProperty "ServerHTTPRequest", true
xmlDoc.load(URL)

If (xmlDoc.parseError.errorCode <> 0) then
    Response.Write "XML Hatasý: " & xmlDoc.parseError.reason
Else

    set channelNodes = xmlDoc.selectNodes("//item/*")
    for each entry in channelNodes
        if entry.tagName = "title" then
                strBaslik = entry.text
		elseif entry.tagName = "link" then
		strLink = entry.text 
		elseif entry.tagName = "description" then
		strBilgi = entry.text        
	Response.Write "&nbsp;<b>" & strBaslik & ": </b>"& strBilgi & " <a target=""_blank"" href="&strLink&">Ýndir >></a><br><br>"
	end if
    next
end If

%></td>