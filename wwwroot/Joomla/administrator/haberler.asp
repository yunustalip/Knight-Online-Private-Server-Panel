<%
'      JoomlASP Site Y�netimi Sistemi (CMS)
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
'      this library; if not, write to the JoomlASP Asp Yaz�l�m Sistemleri., Kargaz Do�al Gaz Bilgi ��lem M�d�rl���
'       36100 Kars / Merkez 
'		Tel : 0544 275 9804 - 0537 275 3655
'		Mail: emre06@hotmail.com.tr / info@joomlasp.com/.net/.org
'
'
'		Lisans Anla�mas� Gere�i L�tfen Google Reklam B�l�m�n� Sitenizden kald�rmay�n�z. Bu sizin GOOGLE reklamlar�n� yapman�za
'		kesinlikle bir engel de�ildir. reklam.asp i�eri�inin yada yay�nlad��� verinin de�i�mesi lisans politikas�n�n d���na ��k�lmas�na
'		ve JoomlASP CMS sistemini �cretsiz yay�nlamak yerine �cretlie hale getirmeye bizi te�fik etmektedir. Bu Sistem i�in verilen eme�e
'		sayg� ve bir �e�it �deme se�ene�i olarak GOOGLE reklam�m�z�n de�i�tirmemesi yada silinmemesi gerekmektedir.
%>
<td height="25" colspan="3" bgcolor="#333333"><span class="style2">&nbsp;JoomlASP'den Haberler </span></td>
        </tr>
      <tr>
        <td height="20" bgcolor="#666666"><span class="style3"> &nbsp;Yeni Mod�ller </span></td>
        <td height="20" bgcolor="#666666"><span class="style3"> &nbsp;Yeni Bile�enler </span></td>
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
    Response.Write "XML Hatas�: " & xmlDoc.parseError.reason
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
    Response.Write "XML Hatas�: " & xmlDoc.parseError.reason
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
        <td valign="top" bgcolor="#999999">&nbsp;<b>Mevcut S�r�m:</b> <%=surum%><br />
          <%
URL = "http://www.joomlasp.com/haberler/rss.asp?istek=versiyon"

set xmlDoc = createObject("MSXML.DOMDocument")
xmlDoc.async = false
xmlDoc.setProperty "ServerHTTPRequest", true
xmlDoc.load(URL)

If (xmlDoc.parseError.errorCode <> 0) then
    Response.Write "XML Hatas�: " & xmlDoc.parseError.reason
Else

    set channelNodes = xmlDoc.selectNodes("//item/*")
    for each entry in channelNodes
        if entry.tagName = "title" then
                strBaslik = entry.text
		elseif entry.tagName = "link" then
		strLink = entry.text 
		elseif entry.tagName = "description" then
		strBilgi = entry.text        
	Response.Write "&nbsp;<b>" & strBaslik & ": </b>"& strBilgi & " <a target=""_blank"" href="&strLink&">�ndir >></a><br><br>"
	end if
    next
end If

%></td>