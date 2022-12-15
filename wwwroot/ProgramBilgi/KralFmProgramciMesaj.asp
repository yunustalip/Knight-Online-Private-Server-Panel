<%response.charset="utf-8"%><style>
.dj_title {
color : #1050B1; 
font-family : "lucida grande", tahoma, verdana, arial, sans-serif;
font-size : 17px; 
text-decoration : none;
} 
.mtn1 {
color : #333333; 
font-family : tahoma, verdana, arial, sans-serif;
font-size : 11px; 
letter-spacing : -0.5px;
text-decoration : none;
} 

</style>
<table width="630" border="0" cellpadding="10" cellspacing="3" background="images/content_bg.png" bgcolor="#FFFFFF">
              <tr> 
                <td valign="top">
				
				    
				
				<form action="http://www.kralfm.com.tr/iletisim.asp?t=studyomesaj" name="form" method="post">
				   <table width="600" border="0" cellspacing="2" cellpadding="2">

                    <tr> 
                      <td><strong><font class="dj_title">Programcılara Mesaj</font></strong><br /><br /><center><font class="mtn1"><font style="color:#0000FF">Bu form aracılığı ile göndereceğiniz mesaj programcıya özel olarak iletilecektir.<br />Canlı yayında okunmayacaktır.<br />Canlı yayına mesaj göndermek istiyorsanız <a href="http://www.kralfm.com.tr/iletisim.asp?t=form" class="mtn1"><font style="color:#AA0000;text-decoration:underline;">tıklayınız.</font></a></font></font></center><br /></td></tr></table>
					  
					  
					  
				  <table width="600" border="0" cellspacing="2" cellpadding="2">
                    <tr> 
                      <td width="190" align="right" class="mtn1">Mesaj Göndermek istediğiniz programcı</td>
                      <td width="10" align="center" class="mtn1">:</td>

                      <td><select name="djname">
					  <option value="">Lütfen Seçiniz</option>
					  
					  <option value="bedirhan@kralfm.com.tr">Bedirhan Gökçe</option>

					  <option value="engin@kralfm.com.tr">Engin Doymuş</option>

					  <option value="harbikiz@kralfm.com.tr">Harbi Kız</option>

					  <option value="kadirhan@kralfm.com.tr">Kadirhan</option>

					  <option value="kahraman@kralfm.com.tr">Kahraman Tazeoğlu</option>

					  <option value="kezban@kralfm.com.tr">Kezban Yaşamul</option>

					  <option value="gezegen@kralfm.com.tr">Mehmet Akbay</option>

					  <option value="melis@kralfm.com.tr">Melis Akçay</option>

					  <option value="serdem@kralfm.com.tr">Serdem Coşkun</option>

					  <option value="sebnem@kralfm.com.tr">Şebnem Öz Doğan</option>

					  </select></td>
                    </tr>
                  </table>
				  
				
				  
				
                  <table width="600" border="0" cellspacing="2" cellpadding="2">

                    <tr> 
                      <td width="190" align="right" class="mtn1">Adınız Soyadınız</td>
                      <td width="10" align="center" class="mtn1">:</td>
                      <td><input name="isim" type="text" size="30"></td>
                    </tr>
                  
                    <tr> 
                      <td align="right" class="mtn1">e-posta</td>
                      <td align="center" class="mtn1">:</td>

                      <td><input name="eposta" type="text" size="30"> <font class="mtn1">( gerekli )</font> </td>
                    </tr>
                  
                      <tr> 
                        <td align="right" class="mtn1">Telefon</td>
                        <td align="center" class="mtn1">:</td>
                        <td><input name="telefon" type="text" size="30"></td>
                      </tr>

                  
                    <tr> 
                      <td align="right" class="mtn1">mesleğiniz</td>
                      <td align="center" class="mtn1">:</td>
                      <td><input name="meslek" type="text" size="30"></td>
                    </tr>
                  
                    <tr> 
                      <td align="right" class="mtn1">şehir</td>
                      <td align="center" class="mtn1">:</td>

                      <td><select onChange="DisBox()" name="sehir">
                          <option selected value=""> - - Lütfen Seçiniz - - </option>
                          <option value=istanbul>İstanbul</option>
                          <option value=Ankara>Ankara</option>
                          <option value=İzmir>İzmir</option>
                          <option value=Adana>Adana</option>

                          <option value=Adıyaman>Adıyaman</option>
                          <option value=Afyon>Afyon</option>
                          <option value=Ağrı>Ağrı</option>
                          <option value=Aksaray>Aksaray</option>
                          <option value=Amasya>Amasya</option>
                          <option value=Antalya>Antalya</option>

                          <option value=Ardahan>Ardahan</option>
                          <option value=Artvin>Artvin</option>
                          <option value=Aydın>Aydın</option>
                          <option value=Balıkesir>Balıkesir</option>
                          <option value=Bartın>Bartın</option>
                          <option value=Batman>Batman</option>

                          <option value=Bayburt>Bayburt</option>
                          <option value=Bilecik>Bilecik</option>
                          <option value=Bingol>Bingol</option>
                          <option value=Bitlis>Bitlis</option>
                          <option value=Bolu>Bolu</option>
                          <option value=Burdur>Burdur</option>

                          <option value=Bursa>Bursa</option>
                          <option value=Çanakkale>Çanakkale</option>
                          <option value=Çankırı>Çankırı</option>
                          <option value=Çorum>Çorum</option>
                          <option value=Denizli>Denizli</option>
                          <option value=Diyarbakır>Diyarbakır</option>

                          <option value=Edirne>Edirne</option>
                          <option value=Elazığ>Elazığ</option>
                          <option value=Erzincan>Erzincan</option>
                          <option value=Erzurum>Erzurum</option>
                          <option value=Eskişehir>Eskişehir</option>
                          <option value=Gaziantep>Gaziantep</option>

                          <option value=Giresun>Giresun</option>
                          <option value=Gümüşhane>Gümüşhane</option>
                          <option value=Hakkari>Hakkari</option>
                          <option value=Hatay>Hatay</option>
                          <option value=Iğdır>Iğdır</option>
                          <option value=Isparta>Isparta</option>

                          <option value=İçel>İçel</option>
                          <option value=Kahramanmaras>Kahramanmaras</option>
                          <option value=Karabük>Karabük</option>
                          <option value=Karaman>Karaman</option>
                          <option value=Kars>Kars</option>
                          <option value=Kastamonu>Kastamonu</option>

                          <option value=Kayseri>Kayseri</option>
                          <option value=Kırıkkale>Kırıkkale</option>
                          <option value=Kırklareli>Kırklareli</option>
                          <option value=Kırşehir>Kırşehir</option>
                          <option value=Kilis>Kilis</option>
                          <option value=Kocaeli>Kocaeli</option>

                          <option value=Konya>Konya</option>
                          <option value=Kütahya>Kütahya</option>
                          <option value=Malatya>Malatya</option>
                          <option value=Manisa>Manisa</option>
                          <option value=Mardin>Mardin</option>
                          <option value=Muğla>Muğla</option>

                          <option value=Muş>Muş</option>
                          <option value=Nevşehir>Nevşehir</option>
                          <option value=Niğde>Niğde</option>
                          <option value=Ordu>Ordu</option>
                          <option value=Osmaniye>Osmaniye</option>
                          <option value=Rize>Rize</option>

                          <option value=Sakarya>Sakarya</option>
                          <option value=Samsun>Samsun</option>
                          <option value=Siirt>Siirt</option>
                          <option value=Sinop>Sinop</option>
                          <option value=Sivas>Sivas</option>
                          <option value=Şanlıurfa>Şanlıurfa</option>

                          <option value=Şırnak>Şırnak</option>
                          <option value=Tekirdağ>Tekirdağ</option>
                          <option value=Tokat>Tokat</option>
                          <option value=Trabzon>Trabzon</option>
                          <option value=Tunceli>Tunceli</option>
                          <option value=Uşak>Uşak</option>

                          <option value=Van>Van</option>
                          <option value=Yalova>Yalova</option>
                          <option value=Yozgat>Yozgat</option>
                          <option value=Zonguldak>Zonguldak</option>
                          <option value=Diger>Diğer...</option>
                        </select></td>

                    </tr>
                  
                      <tr> 
                        <td align="right" class="mtn1">ülke</td>
                        <td align="center" class="mtn1">:</td>
                        <td><select onChange="DisBoxx()" name="ulke">
						<option value="" > - - Lütfen Seçiniz - - </option>
<option selected value="Turkiye">Türkiye</option>

<option value="ABD">ABD</option>
<option value="Afganistan">Afganistan</option>
<option value="Almanya">Almanya</option>
<option value="Andorra">Andorra</option>
<option value="Angola">Angola</option>
<option value="Antarktika">Antarktika</option>
<option value="Arjantin">Arjantin</option>
<option value="Arnavutluk">Arnavutluk</option>
<option value="Avustralya">Avustralya</option>

<option value="Avusturya">Avusturya</option>
<option value="Azerbaycan">Azerbaycan</option>
<option value="Bahama Adaları">Bahama Adaları</option>
<option value="Bahreyn">Bahreyn</option>
<option value="Bangladeş">Bangladeş</option>
<option value="Barbados">Barbados</option>
<option value="Batı Samoa">Batı Samoa</option>
<option value="Belçika">Belçika</option>
<option value="Belize">Belize</option>

<option value="Benin">Benin</option>
<option value="Bermuda">Bermuda</option>
<option value="Beyaz Rusya">Beyaz Rusya</option>
<option value="Bhutan">Bhutan</option>
<option value="B. Arap Emrlk.">B. Arap Emrlk.</option>
<option value="Bolivya">Bolivya</option>
<option value="Bosna Hersek">Bosna Hersek</option>
<option value="Botswana">Botswana</option>
<option value="Brezilya">Brezilya</option>

<option value="Brunei">Brunei</option>
<option value="Bulgaristan">Bulgaristan</option>
<option value="Burkina Faso">Burkina Faso</option>
<option value="Burundi">Burundi</option>
<option value="Cape Verd">Cape Verde</option>
<option value="Cezayir">Cezayir</option>
<option value="Cibuti">Cibuti</option>
<option value="Çad">Çad</option>
<option value="Çek Cum.">Çek Cum.</option>

<option value="Çin">Çin</option>
<option value="Danimarka">Danimarka</option>
<option value="Dominik Cum.">Dominik Cum.</option>
<option value="Dominika">Dominika</option>
<option value="Ekvador">Ekvador</option>
<option value="Ekvator Ginesi">Ekvator Ginesi</option>
<option value="El Salvador">El Salvador</option>
<option value="Eritre">Eritre</option>
<option value="Ermenistan">Ermenistan</option>

<option value="Estonya">Estonya</option>
<option value="Etiyopya">Etiyopya</option>
<option value="Falkland Adaları">Falkland Adaları</option>
<option value="Faroe Adaları">Faroe Adaları</option>
<option value="Fas">Fas</option>
<option value="Fiji">Fiji</option>
<option value="Fildişi Kıyısı">Fildişi Kıyısı</option>
<option value="Filipinler">Filipinler</option>
<option value="Finlandiya">Finlandiya</option>

<option value="Fransa">Fransa</option>
<option value="Gabon">Gabon</option>
<option value="Gambiya">Gambiya</option>
<option value="Gana">Gana</option>
<option value="Gine">Gine</option>
<option value="Gine-Bissau">Gine-Bissau</option>
<option value="Grenada">Grenada</option>
<option value="Grönland">Grönland</option>
<option value="Guatemala">Guatemala</option>

<option value="Guyana">Guyana</option>
<option value="Guney Afrika">Güney Afrika</option>
<option value="Güney Kıbrıs RY">Güney Kıbrıs RY</option>
<option value="Gurcistan">Gürcistan</option>
<option value="Haiti">Haiti</option>
<option value="Hirvatistan">Hırvatistan</option>
<option value="Hindistan">Hindistan</option>
<option value="Hollanda">Hollanda</option>
<option value="Honduras">Honduras</option>

<option value="Irak">Irak</option>
<option value="Indonezya">İndonezya</option>
<option value="Ingiltere">İngiltere</option>
<option value="Iran">İran</option>
<option value="Irlanda">İrlanda</option>
<option value="Ispanya">İspanya</option>
<option value="İsrail">İsrail</option>
<option value="Isveç">İsveç</option>
<option value="Isviçre">İsviçre</option>

<option value="Italya">İtalya</option>
<option value="Izlanda">İzlanda</option>
<option value="Jamaika">Jamaika</option>
<option value="Japonya">Japonya</option>
<option value="Kamboçya">Kamboçya</option>
<option value="Kamerun">Kamerun</option>
<option value="Kanada">Kanada</option>
<option value="Katar">Katar</option>
<option value="Kazakistan">Kazakistan</option>

<option value="Kenya">Kenya</option>
<option value="Kırgızistan">Kırgızistan</option>
<option value="Kiribati">Kiribati</option>
<option value="Kolombiya">Kolombiya</option>
<option value="Komorlar">Komorlar</option>
<option value="Kongo">Kongo</option>
<option value="Kore Güney">Kore Güney</option>
<option value="Kore Kuzey">Kore Kuzey</option>
<option value="Kosta Rika">Kosta Rika</option>

<option value="Kuveyt">Kuveyt</option>
<option value="Kuzey Kıbrıs TC">Kuzey Kıbrıs TC</option>
<option value="Kuba">Küba</option>
<option value="Laos">Laos</option>
<option value="Lesotho">Lesotho</option>
<option value="Letonya">Letonya</option>
<option value="Liberya">Liberya</option>
<option value="Libya">Libya</option>
<option value="Liechtenstein">Liechtenstein</option>

<option value="Litvanya">Litvanya</option>
<option value="Lubnan">Lübnan</option>
<option value="Luksemburg">Lüksemburg</option>
<option value="Macaristan">Macaristan</option>
<option value="Madagaskar">Madagaskar</option>
<option value="Makao">Makao</option>
<option value="Makedonya">Makedonya</option>
<option value="Malavi">Malavi</option>
<option value="Maldiv Adaları">Maldiv Adaları</option>

<option value="Malezya">Malezya</option>
<option value="Mali">Mali</option>
<option value="Malta">Malta</option>
<option value="Mauritius">Mauritius</option>
<option value="Meksika">Meksika</option>
<option value="Mısır">Mısır</option>
<option value="Mogolistan">Moğolistan</option>
<option value="Moldavya">Moldavya</option>
<option value="Monako">Monako</option>

<option value="Moritanya">Moritanya</option>
<option value="Mozambik">Mozambik</option>
<option value="Myanmar">Myanmar</option>
<option value="Namibia">Namibia</option>
<option value="Nauru">Nauru</option>
<option value="Nepal">Nepal</option>
<option value="Nijer">Nijer</option>
<option value="Nijerya">Nijerya</option>
<option value="Nikaragua">Nikaragua</option>

<option value="Norfolk Adası">Norfolk Adası</option>
<option value="Norveç">Norveç</option>
<option value="Ozbekistan">Özbekistan</option>
<option value="Pakistan">Pakistan</option>
<option value="Palau Adaları">Palau Adaları</option>
<option value="Panama">Panama</option>
<option value="Papua-Yeni Gine">Papua-Yeni Gine</option>
<option value="Paraguay">Paraguay</option>
<option value="Peru">Peru</option>

<option value="Polonya">Polonya</option>
<option value="Portekiz">Portekiz</option>
<option value="Puerto Rico">Puerto Rico</option>
<option value="Romanya">Romanya</option>
<option value="Ruanda">Ruanda</option>
<option value="Rusya Fed.">Rusya Fed.</option>
<option value="San Marino">San Marino</option>
<option value="Santa Lucia">Santa Lucia</option>
<option value="Sao Tome">Sao Tome</option>

<option value="Senegal">Senegal</option>
<option value="Seyşeller">Seyşeller</option>
<option value="Sierra Leone">Sierra Leone</option>
<option value="Singapur">Singapur</option>
<option value="Slovakya">Slovakya</option>
<option value="Slovenya">Slovenya</option>
<option value="Solomon Adaları">Solomon Adaları</option>
<option value="Somali">Somali</option>
<option value="Sri Lanka">Sri Lanka</option>

<option value="Sudan">Sudan</option>
<option value="Surinam">Surinam</option>
<option value="Suriye">Suriye</option>
<option value="Suudi Arabistan">Suudi Arabistan</option>
<option value="Svaziland">Svaziland</option>
<option value="Sili">Şili</option>
<option value="Tacikistan">Tacikistan</option>
<option value="Tanzanya">Tanzanya</option>
<option value="Tayland">Tayland</option>

<option value="Tayvan">Tayvan</option>
<option value="Togo">Togo</option>
<option value="Tonga">Tonga</option>
<option value="Tunus">Tunus</option>
<option value="Türkmenistan">Türkmenistan</option>
<option value="Uganda">Uganda</option>
<option value="Ukrayna">Ukrayna</option>
<option value="Umman">Umman</option>
<option value="Uruguay">Uruguay</option>

<option value="Ürdün">Ürdün</option>
<option value="Vanuatu">Vanuatu</option>
<option value="Vatikan">Vatikan</option>
<option value="Venezuela">Venezuela</option>
<option value="Vietnam">Vietnam</option>
<option value="Yemen">Yemen</option>
<option value="Yeni Kaledonya">Yeni Kaledonya</option>
<option value="Yeni Zelanda">Yeni Zelanda</option>
<option value="Yugoslavya">Yugoslavya</option>

<option value="Yunanistan">Yunanistan</option>
<option value="Zambiya">Zambiya</option>
<option value="Zimbabve">Zimbabve</option></select></td>
                      </tr>
                    
                    <tr> 
                      <td align="right" class="mtn1">yaşınız</td>
                      <td align="center" class="mtn1">:</td>
                      <td><input name="yas" type="text" size="10"></td>

                    </tr>
                  
                    <tr> 
                      <td align="right" valign="top" class="mtn1">mesajınız</td>
                      <td align="center" valign="top" class="mtn1">:</td>
                      <td><textarea name="mesaj" cols="24" rows="5"></textarea></td>
                    </tr>
                  
                    <tr> 
                      <td align="right" class="mtn1">&nbsp;</td>

                      <td align="center" class="mtn1">&nbsp;</td>
                      <td><input type="image" src="http://www.kralfm.com.tr/images/btn_gonder.png" name="Submit"></td>
                    </tr>
                  </table>				
                  <table class="noktanokta" width="600" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td>&nbsp;</td>
                    </tr>
                  </table></form>
