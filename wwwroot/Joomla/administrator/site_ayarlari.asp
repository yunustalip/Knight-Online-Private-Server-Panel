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
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;<img src="../images/genel.png" width="128" height="128" align="middle" /><span class="style6"> Genel Ayarlar</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr>
            <td height="20"><%
SQL ="SELECT * FROM gop_ayarlar where ayar_id = '"& 1 &"';"
set rs = server.createobject("ADODB.Recordset")
rs.open SQL , Baglanti
%><form id="form1" name="form1" method="post" action="site_ayarlari.asp?islem=guncelle"><table width="100%" border="0" cellspacing="3" cellpadding="3">
                <tr>
                  <td width="17%"><strong>Site Adý</strong></td>
                  <td width="1%"><strong>:</strong></td>
                  <td width="82%">
                    <input name="siteadi" type="text" class="inputbox" id="siteadi" value="<%=rs("siteadi")%>" size="75" maxlength="75" />                  </td>
                </tr>
                <tr>
                  <td><strong>Site Adresi</strong></td>
                  <td><strong>:</strong></td>
                  <td><input name="siteadres" type="text" class="inputbox" id="siteadres" value="<%=rs("siteadres")%>" size="75" maxlength="75" /> 
                    <span class="style7">Örn: http://www.joomlasp.com/site</span> </td>
                </tr>
                <tr>
                  <td><strong>Gösterilecek Mesaj Sayýsý</strong></td>
                  <td><strong>:</strong></td>
                  <td><input name="vsayi" type="text" class="inputbox" id="vsayi" value="<%=rs("vsayi")%>" size="5" maxlength="3" /> 
                    <span class="style7">varsayýlan: 10</span></td>
                </tr>
                <tr>
                  <td><strong>Mesaj Sütun Sayýsý</strong></td>
                  <td><strong>:</strong></td>
                  <td><input name="vsutun" type="text" class="inputbox" id="vsutun" value="<%=rs("vsutun")%>" size="5" maxlength="2" /> 
                    <span class="style7">varsayýlan: 2</span></td>
                </tr>
                <tr>
                  <td><strong>Mesaj Karakter Sayýsý</strong></td>
                  <td><strong>:</strong></td>
                  <td><input name="vkarakter" type="text" class="inputbox" id="vkarakter" value="<%=rs("vkarakter")%>" size="5" maxlength="3" />                    
                     <span class="style7">varsayýlan: 200</span></td>
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
                    varsayýlan: Ana Sayfa</td>
                </tr>
                <tr>
                  <td><strong>Ana Sayfada Küçük Resim</strong></td>
                  <td>:</td>
                  <td valign="middle" class="style7"><select name="kresim" class="inputbox" id="kresim">
                      <% if rs("kresim") = "1" then %>
                    <option value="<%=rs("kresim")%>" selected="selected">Seçim Göster</option>
                  <% else %>
                    <option value="<%=rs("kresim")%>" selected="selected">Seçim Gösterme</option>
                  <% end if %>
                      <option value="1">Göster</option>
                      <option value="0">Gösterme</option>
                    </select> 
                  varsayýlan: Göster</td>
                </tr>
                <tr>
                  <td><strong>AspJpeg Kullan</strong></td>
                  <td><strong>:</strong></td>
                  <td valign="middle"><select name="aspjpeg" class="inputbox" id="aspjpeg">
                  <% if rs("aspjpeg") = "evet" then %>
                    <option value="<%=rs("aspjpeg")%>" selected="selected">Seçim Evet</option>
                  <% else %>
                    <option value="<%=rs("aspjpeg")%>" selected="selected">Seçim Hayýr</option>
                  <% end if %>
                    <option value="evet">Evet</option>
                    <option value="hayir">Hayýr</option>
                    
                                                                        </select>
                    <span class="style7">varsayýlan: Evet (Lütfen sunucunuzun AspJpeg destekleyip desteklemediðini öðreniniz.)</span></td>
                </tr>
                <tr>
                  <td><strong>Sitede Yorumlarý Göster</strong></td>
                  <td><strong>:</strong></td>
                  <td valign="middle">
                    <select name="vyorum" class="inputbox" id="vyorum">
                  <% if rs("vyorum") = "goster" then %>
                    <option value="<%=rs("vyorum")%>" selected="selected">Seçim Göster</option>
                  <% else %>
                    <option value="<%=rs("vyorum")%>" selected="selected">Seçim Gösterme</option>
                  <% end if %>
                      <option value="goster">Göster</option>
                      <option value="gosterme">Gösterme</option>
                    </select>                    
                     <span class="style7">varsayýlan: Göster                  </span></td>
                </tr>
                <tr>
                  <td><strong>Google Reklamlarý</strong></td>
                  <td>&nbsp;</td>
                  <td><select name="google" class="inputbox" id="google">
                    <% if rs("google") = "1" then %>
                    <option value="<%=rs("google")%>" selected="selected">Seçim Göster</option>
                    <% else %>
                    <option value="<%=rs("google")%>" selected="selected">Seçim Gösterme</option>
                    <% end if %>
                    <option value="1">Göster</option>
                    <option value="0">Gösterme</option>
                  </select> <span class="style7">varsayýlan: Göster </span></td>
                </tr>
                <tr>
                  <td><strong>Site Yenileme</strong></td>
                  <td>&nbsp;</td>
                  <td><input name="yenile" type="text" id="yenile" value="<%=rs("yenileme")%>" size="3" maxlength="3" />
                    <span class="style7">varsayýlan: 240 </span></td>
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
Response.Write "Yüklü dil bulunamadý"
else
Response.Write "<option value="""&lang("lang_id")&""">"&lang("lang_adi")&"</option>"
end if
%>
                    </select>
                    varsayýlan: Türkçe</td>
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
