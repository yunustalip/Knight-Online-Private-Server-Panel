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
<!--#include file="admin_a.asp"-->
<%
islem = request.querystring("islem")
if islem = "guncelle" then
call guncelle
elseif islem = "" then
call default
end if
sub default
%>
<table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;<img src="../images/anket.png" width="128" height="128" align="middle" /><span class="style6"> Ýstatistlik Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr>
            <td height="20"><table width="100%" border="0" cellpadding="2" cellspacing="2" bgcolor="#FFFFFF">
              <tr>
                <td width="40%"><%
set rs = baglanti.execute("SELECT * FROM gop_sayacayar;")
%><form id="form1" name="form1" method="post" action="sayac_ayarlari.asp?islem=guncelle">
                  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
                <tr>
                  <td width="170" height="30" bgcolor="fbe8a6"><span style="font-weight: bold"> &nbsp;Online Kullanýcý Sayýsý</span></td>
                  <td width="1%" height="25" bgcolor="fbe8a6"><span style="font-weight: bold">:</span></td>
                  <td height="25" valign="middle" bgcolor="fbe8a6" class="style7">
                    <div align="center">
                      <% if rs("online") = "1" then %>
                      Göster 
  <input name="online" type="radio" id="radio" value="1" checked="checked">
 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                    Gösterme 
  <input type="radio" name="online" id="radio2" value="0" >
  <% else %>
                      Göster 
  <input type="radio" name="online" id="radio" value="1" />
   &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                    Gösterme 
  <input type="radio" name="online" id="radio2" value="0" checked="checked"/>
  <% end if %>
                    </div></td>
                </tr>
                <tr>
                  <td height="30"><span style="font-weight: bold">&nbsp;Bugünkü Tekil Ziyaretçi</span></td>
                  <td height="25"><span style="font-weight: bold">:</span></td>
                  <td height="25" valign="middle">
                    <div align="center">
                      <% if rs("btekil") = "1" then %>
                      Göster 
  <input name="btekil" type="radio" value="1" checked="checked">
   &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                    Gösterme 
  <input name="btekil" type="radio" value="0" >
  <% else %>
                      Göster 
  <input type="radio" name="btekil" value="1" />
   &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                    Gösterme 
  <input type="radio" name="btekil" value="0" checked="checked"/>
  <% end if %>
                    </div></td>
                </tr>
                <tr>
                  <td height="30" bgcolor="fbe8a6"><span style="font-weight: bold">&nbsp;Bugünkü Çoðul Ziyaretçi</span></td>
                  <td height="25" bgcolor="fbe8a6"><span style="font-weight: bold">:</span></td>
                  <td height="25" valign="middle" bgcolor="fbe8a6">
                    <div align="center">
                      <% if rs("bcogul") = "1" then %>
                      Göster 
  <input name="bcogul" type="radio" value="1" checked="checked">
   &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                    Gösterme 
  <input name="bcogul" type="radio" value="0" >
  <% else %>
                      Göster 
  <input type="radio" name="bcogul" value="1" />
   &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                    Gösterme 
  <input type="radio" name="bcogul" value="0" checked="checked"/>
  <% end if %>
                    </div></td>
                </tr>
                <tr>
                  <td height="30"><span style="font-weight: bold">&nbsp;Dünkü Tekil Ziyaretçi</span></td>
                  <td height="25"><span style="font-weight: bold">:</span></td>
                  <td height="25">
                    <div align="center">
                      <% if rs("dtekil") = "1" then %>
                      Göster 
  <input name="dtekil" type="radio" value="1" checked="checked">
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                   Gösterme 
  <input name="dtekil" type="radio" value="0" >
  <% else %>
                      Göster 
  <input type="radio" name="dtekil" value="1" />
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                   Gösterme 
  <input type="radio" name="dtekil" value="0" checked="checked"/>
  <% end if %> 
                    </div></td>
                </tr>
                <tr>
                  <td height="30" bgcolor="fbe8a6"><span style="font-weight: bold">&nbsp;Dünkü Çoðul Ziyaretçi</span></td>
                  <td height="25" bgcolor="fbe8a6"><span style="font-weight: bold">:</span></td>
                  <td height="25" bgcolor="fbe8a6">
                    <div align="center">
                      <% if rs("dcogul") = "1" then %>
                      Göster 
  <input name="dcogul" type="radio" value="1" checked="checked">
   &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                    Gösterme 
  <input name="dcogul" type="radio" value="0" >
  <% else %>
                      Göster 
  <input type="radio" name="dcogul" value="1" />
     &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                  Gösterme 
  <input type="radio" name="dcogul" value="0" checked="checked"/>
  <% end if %> 
                    </div></td>
                </tr>
                <tr>
                  <td height="30"><span style="font-weight: bold">&nbsp;Toplam Tekil Ziyaretçi</span></td>
                  <td height="25"><span style="font-weight: bold">:</span></td>
                  <td height="25">

                    <div align="center">
                      <% if rs("toplamt") = "1" then %>
                      Göster 
  <input name="toplamt" type="radio" value="1" checked="checked">
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                 Gösterme 
  <input name="toplamt" type="radio" value="0" >
  <% else %>
                      Göster 
  <input type="radio" name="toplamt" value="1" />
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                 Gösterme 
  <input type="radio" name="toplamt" value="0" checked="checked"/>
  <% end if %> 
                    </div></td>
                </tr>
                <tr>
                  <td height="30" bgcolor="fbe8a6"><span style="font-weight: bold">&nbsp;Toplam Çoðul Ziyaretçi</span></td>
                  <td height="25" bgcolor="fbe8a6"><span style="font-weight: bold">:</span></td>
                  <td height="25" bgcolor="fbe8a6">
                    <div align="center">
                      <% if rs("toplamc") = "1" then %>
                      Göster 
  <input name="toplamc" type="radio" value="1" checked="checked">
     &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                  Gösterme 
  <input name="toplamc" type="radio" value="0" >
  <% else %>
                      Göster 
  <input type="radio" name="toplamc" value="1" />
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                 Gösterme 
  <input type="radio" name="toplamc" value="0" checked="checked"/>
  <% end if %>
                    </div></td>
                </tr>
                <tr>
                  <td height="30"><span style="font-weight: bold">&nbsp;Aktif Ziyaretçi</span></td>
                  <td height="25"><span style="font-weight: bold">:</span></td>
                  <td height="25">
                    <div align="center">
                      <% if rs("aktifuye") = "1" then %>
                      Göster 
  <input name="aktifuye" type="radio" value="1" checked="checked">
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                 Gösterme 
  <input name="aktifuye" type="radio" value="0" >
  <% else %>
                      Göster 
  <input type="radio" name="aktifuye" value="1" />
       &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                Gösterme 
  <input type="radio" name="aktifuye" value="0" checked="checked"/>
  <% end if %>
                    </div></td>
                </tr>
                                <tr>
                                  <td height="30" bgcolor="fbe8a6"><span style="font-weight: bold">&nbsp;Son Üye</span></td>
                                  <td height="25" bgcolor="fbe8a6"><span style="font-weight: bold">:</span></td>
                                  <td height="25" bgcolor="fbe8a6">
                                    <div align="center">
                                      <% if rs("sonuye") = "1" then %>
Göster 
<input name="sonuye" type="radio" value="1" checked="checked">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Gösterme 
<input name="sonuye" type="radio" value="0" >
<% else %>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Göster 
<input type="radio" name="sonuye" value="1" />
Gösterme 
<input type="radio" name="sonuye" value="0" checked="checked"/>
<% end if %>
                                    </div></td>
                            </tr>
                                <tr>
                                  <td height="30"><span style="font-weight: bold">&nbsp;Veri Sayýsý</span></td>
                                  <td height="25"><span style="font-weight: bold">:</span></td>
                                  <td height="25">
                                    <div align="center">
                                      <% if rs("veri") = "1" then %>
Göster 
<input name="veri" type="radio" value="1" checked="checked">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Gösterme 
<input name="veri" type="radio" value="0" >
<% else %>
Göster 
<input type="radio" name="veri" value="1" />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Gösterme 
<input type="radio" name="veri" value="0" checked="checked"/>
<% end if %>
                                    </div></td>
                            </tr>
                                <tr>
                                  <td height="30" bgcolor="fbe8a6"><span style="font-weight: bold">&nbsp;Okunma Sayýsý</span></td>
                                  <td height="25" bgcolor="fbe8a6"><span style="font-weight: bold">:</span></td>
                                  <td height="25" bgcolor="fbe8a6">
                                    <div align="center">
                                      <% if rs("okunma") = "1" then %>
Göster 
<input name="okunma" type="radio" value="1" checked="checked">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Gösterme 
<input name="okunma" type="radio" value="0" >
<% else %>
Göster 
<input type="radio" name="okunma" value="1" />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Gösterme 
<input type="radio" name="okunma" value="0" checked="checked"/>
<% end if %>
                                    </div></td>
                            </tr>
                                <tr>
                                  <td height="30"><span style="font-weight: bold">&nbsp;IP Adresi</span></td>
                                  <td height="25"><span style="font-weight: bold">:</span></td>
                                  <td height="25">
                                    <div align="center">
                                      <% if rs("ip") = "1" then %>
Göster 
<input name="ip" type="radio" value="1" checked="checked">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Gösterme 
<input name="ip" type="radio" value="0" >
<% else %>
Göster 
<input type="radio" name="ip" value="1" />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Gösterme 
<input type="radio" name="ip" value="0" checked="checked"/>
<% end if %>
                                    </div></td>
                      </tr>
                      <tr>
                  <td height="25">&nbsp;</td>
                  <td height="25">&nbsp;</td>
                  <td height="25">
                    <input name="button" type="submit" class="button" id="button" value="Kaydet" />                  </td>
                </tr>
            </table>
            </form></td>
                <td width="60%" valign="top" bgcolor="#FBE8A6">

                <table width="100%" border="0" style="border:solid 1px; border-color:#000000;">
                  <tr>
                    <td bgcolor="#990000" style="border:solid 1px; border-color:#000000;"><div align="center" class="style4" style="font-weight: bold">Tarih</div></td>
                    <td bgcolor="#990000" style="border:solid 1px; border-color:#000000;"><div align="center" class="style4" style="font-weight: bold">Tekil</div></td>
                    <td bgcolor="#990000" style="border:solid 1px; border-color:#000000;"><div align="center" class="style4" style="font-weight: bold">Çoðul</div></td>
                  </tr>
                  <%
listele = 20
If Request.QueryString("s") <> "" Then 
  Sayfa = CInt(Request.QueryString("s"))
Else 
  Sayfa = 1 
End If 

set ist = baglanti.execute("SELECT * FROM gop_sayac order by say_id desc LIMIT "& (listele*Sayfa)-(listele) & "," & listele )

Set SQLToplam = baglanti.Execute("select count(say_id) from gop_sayac") 
TopKayit = SQLToplam(0)

do while not ist.eof or ist.bof
response.Write ""
%>
                  <tr bgcolor="#FFFFFF">
                    <td><div align="center"><%= ist("sayac_tarih") %></div></td>
                    <td><div align="center"><%= ist("sayac_tekil") %></div></td>
                    <td><div align="center"><%= ist("sayac_cogul") %></div></td>
                  </tr>
<%
ist.movenext
loop
ist.close

Set isttopt = baglanti.Execute("SELECT SUM(sayac_tekil) AS tekil FROM gop_sayac")
Set isttopc = baglanti.Execute("SELECT SUM(sayac_cogul) AS cogul FROM gop_sayac")
%>
<tr bgcolor="#FFCC33">
                    <td><div align="center" style="font-weight: bold">Toplam</div></td>
                    <td><div align="center" style="font-weight: bold"><%= isttopt("tekil") %></div></td>
                    <td><div align="center" style="font-weight: bold"><%= isttopc("cogul") %></div></td>
                  </tr>
                </table><!--#include file="sayfa.asp"--></td>
                </tr>
              
              
            </table></td>
          </tr>
        </table>
        
<%
end sub

sub guncelle
ip = Request.Form("ip")
btekil = Request.Form("btekil")
bcogul = Request.Form("bcogul")
dtekil = Request.Form ("dtekil")
dcogul = Request.Form("dcogul")
toplamt = Request.Form("toplamt")
toplamc = Request.Form("toplamc")
aktifuye = Request.Form("aktifuye")
sonuye = Request.Form("sonuye")
veri = Request.Form("veri")
okunma = Request.Form("okunma")
online = Request.Form("online")


baglanti.Execute("UPDATE gop_sayacayar set ip='"&ip&"', btekil='"&btekil&"', bcogul='"&bcogul&"',dtekil='"&dtekil&"', dcogul='"&dcogul&"', toplamt='"&toplamt&"', toplamc='"&toplamc&"', aktifuye='"&aktifuye&"', okunma='"&okunma&"', online='"&online&"', sonuye='"&sonuye&"', veri='"&veri&"';")
Response.Redirect request.ServerVariables("HTTP_REFERER")
end sub
%>
        <!--#include file="admin_b.asp"-->
