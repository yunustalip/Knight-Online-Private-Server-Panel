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
<!--#include file="../functions/fonksiyonlar.asp"-->
<!--#include file="../functions/filitre.asp"--><head>
<meta http-equiv="Content-Language" content="tr">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-9" />
<meta name="author" content="JoomlASP | Hasan Emre Asker">
<meta name="keywords" content="<%=metakey%>">
<meta name="description" content="<%=metadesc%>">
<%
if yenile <> "0" then
Response.Write "<meta http-equiv=""refresh"" content="""&yenileme&""">"
end if

'Yukardaki b�l�m�kendinzie g�re d�zenleyiniz
%>
<link href="favicon.ico" rel="JoomlASP" />
<link href="styles/joomlasp.css" rel="stylesheet" type="text/css" />
</head>
<script type="text/javascript" src="JavaScript/persistmenu.js"></script>
<script type="text/javascript" src="JavaScript/popup.js"></script>
<script type="text/javascript" src="JavaScript/anket.js"></script>
<script type="text/javascript" src="JavaScript/pageear.js"></script>
<script type="text/javascript" src="JavaScript/AC_RunActiveContent.js"></script>
<script type="text/javascript" src="JavaScript/AC_OETags.js"></script>   
<script type="text/javascript">    
    writeObjects();
</script>
<!--#include file="../includes/sayac.asp"-->