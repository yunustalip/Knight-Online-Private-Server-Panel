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
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;<img src="../images/genel.png" width="128" height="128" align="middle" /><span class="style6"> Genel Ayarlar</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr>
            <td height="20"><%
SQL ="SELECT * FROM gop_ayarlar where ayar_id = '"& 1 &"';"
set rs = server.createobject("ADODB.Recordset")
rs.open SQL , Baglanti
%><form id="form1" name="form1" method="post" action="site_ayarlari.asp?islem=guncelle"><table width="100%" border="0" cellspacing="3" cellpadding="3">
                <tr>
                  <td width="17%"><strong>Site Ad�</strong></td>
                  <td width="1%"><strong>:</strong></td>
                  <td width="82%">
                    <input name="siteadi" type="text" class="inputbox" id="siteadi" value="<%=rs("siteadi")%>" size="75" maxlength="75" />                  </td>
                </tr>
                <tr>
                  <td><strong>Site Adresi</strong></td>
                  <td><strong>:</strong></td>
                  <td><input name="siteadres" type="text" class="inputbox" id="siteadres" value="<%=rs("siteadres")%>" size="75" maxlength="75" /> 
                    <span class="style7">�rn: http://www.joomlasp.com/site</span> </td>
                </tr>
                <tr>
                  <td><strong>G�sterilecek Mesaj Say�s�</strong></td>
                  <td><strong>:</strong></td>
                  <td><input name="vsayi" type="text" class="inputbox" id="vsayi" value="<%=rs("vsayi")%>" size="5" maxlength="3" /> 
                    <span class="style7">varsay�lan: 10</span></td>
                </tr>
                <tr>
                  <td><strong>Mesaj S�tun Say�s�</strong></td>
                  <td><strong>:</strong></td>
                  <td><input name="vsutun" type="text" class="inputbox" id="vsutun" value="<%=rs("vsutun")%>" size="5" maxlength="2" /> 
                    <span class="style7">varsay�lan: 2</span></td>
                </tr>
                <tr>
                  <td><strong>Mesaj Karakter Say�s�</strong></td>
                  <td><strong>:</strong></td>
                  <td><input name="vkarakter" type="text" class="inputbox" id="vkarakter" value="<%=rs("vkarakter")%>" size="5" maxlength="3" />                    
                     <span class="style7">varsay�lan: 200</span></td>
                </tr>
                <tr>
                  <td><strong>Ana Sayfa</strong></td>
                  <td>:</td>
                  <td valign="middle" class="style7">
                  
                  
<select name="eklenti_k" class="inputbox" id="eklenti_k">
<%
SQLekl ="SELECT * from gop_eklentiler;"
set ekl = server.createobject("ADODB.Recordset")
ekl.open SQLekl , Baglanti
if ekl.eof or ekl.bof then 
Response.write " "
else

do while not ekl.eof

Response.Write "<option value="""& ekl("eklenti_k") &""">"& ekl("eklenti_adi") &"</option>"

ekl.movenext
loop


ekl.close
set ekl=nothing
end if
%>
</select>
                    varsay�lan: Ana Sayfa</td>
                </tr>
                <tr>
                  <td><strong>Ana Sayfada K���k Resim</strong></td>
                  <td>:</td>
                  <td valign="middle" class="style7"><select name="kresim" class="inputbox" id="kresim">
                      <% if rs("kresim") = "1" then %>
                    <option value="<%=rs("kresim")%>" selected="selected">Se�im G�ster</option>
                  <% else %>
                    <option value="<%=rs("kresim")%>" selected="selected">Se�im G�sterme</option>
                  <% end if %>
                      <option value="1">G�ster</option>
                      <option value="0">G�sterme</option>
                    </select> 
                  varsay�lan: G�ster</td>
                </tr>
                <tr>
                  <td><strong>AspJpeg Kullan</strong></td>
                  <td><strong>:</strong></td>
                  <td valign="middle"><select name="aspjpeg" class="inputbox" id="aspjpeg">
                  <% if rs("aspjpeg") = "evet" then %>
                    <option value="<%=rs("aspjpeg")%>" selected="selected">Se�im Evet</option>
                  <% else %>
                    <option value="<%=rs("aspjpeg")%>" selected="selected">Se�im Hay�r</option>
                  <% end if %>
                    <option value="evet">Evet</option>
                    <option value="hayir">Hay�r</option>
                    
                                                                        </select>
                    <span class="style7">varsay�lan: Evet (L�tfen sunucunuzun AspJpeg destekleyip desteklemedi�ini ��reniniz.)</span></td>
                </tr>
                <tr>
                  <td><strong>Sitede Yorumlar� G�ster</strong></td>
                  <td><strong>:</strong></td>
                  <td valign="middle">
                    <select name="vyorum" class="inputbox" id="vyorum">
                  <% if rs("vyorum") = "goster" then %>
                    <option value="<%=rs("vyorum")%>" selected="selected">Se�im G�ster</option>
                  <% else %>
                    <option value="<%=rs("vyorum")%>" selected="selected">Se�im G�sterme</option>
                  <% end if %>
                      <option value="goster">G�ster</option>
                      <option value="gosterme">G�sterme</option>
                    </select>                    
                     <span class="style7">varsay�lan: G�ster                  </span></td>
                </tr>
                <tr>
                  <td><strong>Google Reklamlar�</strong></td>
                  <td>&nbsp;</td>
                  <td><select name="google" class="inputbox" id="google">
                    <% if rs("google") = "1" then %>
                    <option value="<%=rs("google")%>" selected="selected">Se�im G�ster</option>
                    <% else %>
                    <option value="<%=rs("google")%>" selected="selected">Se�im G�sterme</option>
                    <% end if %>
                    <option value="1">G�ster</option>
                    <option value="0">G�sterme</option>
                  </select> <span class="style7">varsay�lan: G�ster </span></td>
                </tr>
                <tr>
                  <td><strong>Site Yenileme</strong></td>
                  <td>&nbsp;</td>
                  <td><input name="yenile" type="text" id="yenile" value="<%=rs("yenileme")%>" size="3" maxlength="3" />
                    <span class="style7">varsay�lan: 240 </span></td>
                </tr>
                <tr>
                  <td><strong>Google Adsense Kodu</strong></td>
                  <td>&nbsp;</td>
                  <td><input name="google_code" type="text" class="inputbox" id="google_code" value="<%=rs("google_code")%>" size="75" /></td>
                </tr>
                <tr>
                  <td><strong>Meta Description</strong></td>
                  <td><strong>:</strong></td>
                  <td><textarea name="meta_desc" cols="74" class="inputbox2" id="meta_desc"><%=rs("meta_desc")%></textarea></td>
                </tr>
                <tr>
                  <td><strong>Meta Keywords</strong></td>
                  <td><strong>:</strong></td>
                  <td><textarea name="meta_key" cols="74" class="inputbox2" id="meta_key"><%=rs("meta_key")%></textarea></td>
                </tr>
                <tr>
                  <td><strong>Admin Mail Adresi</strong></td>
                  <td><strong>:</strong></td>
                  <td><input name="admin_mail" type="text" class="inputbox" id="admin_mail" value="<%=rs("admin_mail")%>" size="75" /></td>
                </tr>
                <tr>
                  <td><strong>Site Dili</strong></td>
                  <td>:</td>
                  <td valign="middle" class="style7">
<select name="dilayari" class="inputbox" id="dilayari">
<%
set lang = baglanti.execute("select * from gop_language")
if lang.eof or lang.bof then
Response.Write "Y�kl� dil bulunamad�"
else
Response.Write "<option value="""&lang("lang_id")&""">"&lang("lang_adi")&"</option>"
end if
%>
                    </select>
                    varsay�lan: T�rk�e</td>
                </tr>
                <tr>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                  <td>
                    <input name="button" type="submit" class="button" id="button" value="Kaydet" />                  </td>
                </tr>
            </table>
            </form>
           </td>
          </tr>
        </table>
<%
end sub

sub guncelle
vsayi = Request.Form("vsayi")
vsutun = Request.Form("vsutun")
vkarakter = Request.Form("vkarakter")
vyorum = Request.Form ("vyorum")
siteadi = Request.Form("siteadi")
siteadres = Request.Form("siteadres")
admin_mail = Request.Form("admin_mail")
meta_desc = Request.Form("meta_desc")
meta_key = Request.Form("meta_key")
aspjpeg = Request.Form("aspjpeg")
kresim = Request.Form("kresim")
google = Request.Form("google")
google_code = Request.Form("google_code")
yenile = Request.Form("yenile")
eklenti_k = Request.Form("eklenti_k")
dilayari = Request.Form("dilayari")


baglanti.Execute("UPDATE gop_ayarlar set vsayi='"&vsayi&"', vsutun='"&vsutun&"', vkarakter='"&vkarakter&"',vyorum='"&vyorum&"', siteadi='"&siteadi&"', admin_mail='"&admin_mail&"', meta_desc='"&meta_desc&"', meta_key='"&meta_key&"', aspjpeg='"&aspjpeg&"', siteadres='"&siteadres&"', kresim='"&kresim&"', google='"&google&"', google_code='"&google_code&"', yenileme='"&yenile&"', eklenti_k='"&eklenti_k&"', dilayari='"&dilayari&"' where ayar_id='" & 1 & "' ;")
Response.Redirect request.ServerVariables("HTTP_REFERER")
end sub
%>
        <!--#include file="admin_b.asp"-->
