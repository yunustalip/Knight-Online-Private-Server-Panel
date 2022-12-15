
			<table border="0" width="550" id="table1" cellpadding="0" style="border-collapse: collapse">
				<tr>
					<td height="26" width="15">
					<img border="0" src="tema/images/blok_1.gif" width="15" height="26"></td>
					<td height="26" background="tema/images/blok_2.gif" width="525">
					<p align="center">
					<font class="baslik">Paylaþ</font></td>
					<td height="26" width="10">
					<img border="0" src="tema/images/blok_3.gif" width="10" height="26"></td>
				</tr>
				<tr>
					<td class="o_tab" valign="top" width="550" colspan="3" align="center">
							<table border="0" width="530" id="table1" cellpadding="0" style="border-collapse: collapse">
<%
adres=Cevir(AramaFiltre(Request.ServerVariables("HTTP_REFERER")))
title=Cevir(AramaFiltre(Request.QueryString("baslik")))
if adres="" or title="" then
%>
								<tr>
									<td align="center"><font class="orta">Sayfaya Güvenli Yollardan Girilmedi</font></td>
								</tr>
<% Else %>
								<tr>
									<td colspan="3"><b><font class="orta">Baþlýk: <%=title%><br>Adres: <%=adres%><br></font></b></td>
								</tr>
								<tr>
									<td width="33%" valign="top">
								<a title="Ask" target="_blank" href="http://myjeeves.ask.com/mysearch/BookmarkIt?v=1.2&t=webpages&url=<%=adres%>&title=<%=title%>">
								<img alt="Ask" src="tema/v3_img/paylas/ask.png"border="0"></a>
								<a title="Ask" target="_blank" href="http://myjeeves.ask.com/mysearch/BookmarkIt?v=1.2&t=webpages&url=<%=adres%>&title=<%=title%>">
								Ask</a>
								<br>
								<a title="Backflip" target="_blank" href="http://www.backflip.com/add_page_pop.ihtml?url=<%=adres%>&title=<%=title%>">
								<img alt="Backflip" src="tema/v3_img/paylas/backflip.png"border="0"></a>
								<a title="Backflip" target="_blank" href="http://www.backflip.com/add_page_pop.ihtml?url=<%=adres%>&title=<%=title%>">
								Backflip</a>
								<br>
								<a title="Bað-Kur" target="_blank" href="http://www.bagcik.com/person_links/new?url=<%=adres%>&title=<%=title%>">
								<img alt="Bað-Kur" src="tema/v3_img/paylas/bagcik.gif"border="0"></a>
								<a title="Bað-Kur" target="_blank" href="http://www.bagcik.com/person_links/new?url=<%=adres%>&title=<%=title%>">
								Bað-Kur</a>
								<br>
								<a title="BlinkBits" target="_blank" href="http://www.blinkbits.com/bookmarklets/save.php?v=1&source_url=<%=adres%>&title=<%=title%>&body=TITLE">
								<img alt="BlinkBits" src="tema/v3_img/paylas/blinkbits.png"border="0"></a>
								<a title="BlinkBits" target="_blank" href="http://www.blinkbits.com/bookmarklets/save.php?v=1&source_url=<%=adres%>&title=<%=title%>&body=TITLE">
								BlinkBits</a>
								<br>
								<a title="BlinkList" target="_blank" href="http://www.blinklist.com/index.php?Action=Blink/addblink.php&Description=&Url=<%=adres%>&Title=<%=title%>">
								<img alt="BlinkList" src="tema/v3_img/paylas/blinklist.png"border="0"></a>
								<a title="BlinkList" target="_blank" href="http://www.blinklist.com/index.php?Action=Blink/addblink.php&Description=&Url=<%=adres%>&Title=<%=title%>">
								BlinkList</a>
								<br>
								<a title="Blogmarks" target="_blank" href="http://blogmarks.net/my/new.php?mini=1&simple=1&url=<%=adres%>&title=<%=title%>">
								<img alt="Blogmarks" src="tema/v3_img/paylas/blogmarks.png"border="0"></a>
								<a title="Blogmarks" target="_blank" href="http://blogmarks.net/my/new.php?mini=1&simple=1&url=<%=adres%>&title=<%=title%>">
								Blogmarks</a>
								<br>
								<a title="Blue Dot" target="_blank" href="http://bluedot.us/Authoring.aspx?u=<%=adres%>&t=<%=title%>">
								<img alt="Blue Dot" src="tema/v3_img/paylas/bluedot.png"border="0"></a>
								<a title="Blue Dot" target="_blank" href="http://bluedot.us/Authoring.aspx?u=<%=adres%>&t=<%=title%>">
								Blue Dot</a>
								<br>
								<a title="co.mments" target="_blank" href="http://co.mments.com/track?url=<%=adres%>&title=<%=title%>">
								<img alt="co.mments" src="tema/v3_img/paylas/co.mments.gif"border="0"></a>
								<a title="co.mments" target="_blank" href="http://co.mments.com/track?url=<%=adres%>&title=<%=title%>">
								co.mments</a>
								<br>
								<a title="connotea" target="_blank" href="http://www.connotea.org/addpopup?continue=confirm&uri=<%=adres%>&title=<%=title%>">
								<img alt="connotea" src="tema/v3_img/paylas/connotea.png"border="0"></a>
								<a title="connotea" target="_blank" href="http://www.connotea.org/addpopup?continue=confirm&uri=<%=adres%>&title=<%=title%>">
								connotea</a>
								<br>
								<a title="Del.icio.us" target="_blank" href="http://del.icio.us/post?url=<%=adres%>&title=<%=title%>">
								<img alt="Del.icio.us" src="tema/v3_img/paylas/delicious.png"border="0"></a>
								<a title="Del.icio.us" target="_blank" href="http://del.icio.us/post?url=<%=adres%>&title=<%=title%>">
								Del.icio.us</a>
								<br>
								<a title="Face Book" target="_blank" href="http://www.facebook.com/sharer.php?u=<%=adres%>&t=<%=title%>">
								<img alt="Face Book" src="tema/v3_img/paylas/facebook.gif"border="0"></a>
								<a title="Face Book" target="_blank" href="http://www.facebook.com/sharer.php?u=<%=adres%>&t=<%=title%>">
								FaceBook</a>
								<br>
								<a title="Digg" target="_blank" href="http://digg.com/submit?phase=2&url=<%=adres%>&title=<%=title%>">
								<img alt="Digg" src="tema/v3_img/paylas/digg.png"border="0"></a>
								<a title="Digg" target="_blank" href="http://digg.com/submit?phase=2&url=<%=adres%>&title=<%=title%>">
								Digg</a>
								<br>
								<a title="Fark" target="_blank" href="http://cgi.fark.com/cgi/fark/edit.pl?new_url=<%=adres%>&new_comment=<%=title%>&new_comment=Günlük Haftalýk Aylýk&linktype=Misc">
								<img alt="Fark" src="tema/v3_img/paylas/fark.png"border="0"></a>
								<a title="Fark" target="_blank" href="http://cgi.fark.com/cgi/fark/edit.pl?new_url=<%=adres%>&new_comment=<%=title%>&new_comment=Günlük Haftalýk Aylýk&linktype=Misc">
								Fark</a>
									</td>
									<td width="33%" align="left" valign="top">
								<a title="Limkle!" target="_blank" href="http://www.limk.com/limkle.php?page=addlink&title=<%=title%>&url=<%=adres%>">
								<img alt="Limkle!" src="tema/v3_img/paylas/limk.gif"border="0"></a>
								<a title="Limkle!" target="_blank" href="http://www.limk.com/limkle.php?page=addlink&title=<%=title%>&url=<%=adres%>">
								Limkle!</a>
								<br>
								<a title="LinkaGoGo" target="_blank" href="http://www.linkagogo.com/go/AddNoPopup?url=<%=adres%>&title=<%=title%>">
								<img alt="LinkaGoGo" src="tema/v3_img/paylas/linkagogo.png"border="0"></a>
								<a title="LinkaGoGo" target="_blank" href="http://www.linkagogo.com/go/AddNoPopup?url=<%=adres%>&title=<%=title%>">
								LinkaGoGo</a>
								<br>
								<a title="Live Bookmarks" target="_blank" href="https://favorites.live.com/quickadd.aspx?marklet=1&mkt=en-us&url=<%=adres%>&title=<%=title%>&top=1">
								<img alt="Live Bookmarks" src="tema/v3_img/paylas/windows_live.gif"border="0"></a>
								<a title="Live Bookmarks" target="_blank" href="https://favorites.live.com/quickadd.aspx?marklet=1&mkt=en-us&url=<%=adres%>&title=<%=title%>&top=1">
								Live Bookmarks</a>
								<br>
								<a title="Ma.gnolia" target="_blank" href="http://ma.gnolia.com/beta/bookmarklet/add?url=<%=adres%>&title=<%=title%>&description=TITLE">
								<img alt="Ma.gnolia" src="tema/v3_img/paylas/magnolia.png"border="0"></a>
								<a title="Ma.gnolia" target="_blank" href="http://ma.gnolia.com/beta/bookmarklet/add?url=<%=adres%>&title=<%=title%>&description=TITLE">
								Ma.gnolia</a>
								<br>
								<a title="NewsVine" target="_blank" href="http://www.newsvine.com/_tools/seed&save?u=<%=adres%>&h=<%=title%>">
								<img alt="NewsVine" src="tema/v3_img/paylas/newsvine.png"border="0"></a>
								<a title="NewsVine" target="_blank" href="http://www.newsvine.com/_tools/seed&save?u=<%=adres%>&h=<%=title%>">
								NewsVine</a>
								<br>
								<a title="Netscape" target="_blank" href="http://www.netscape.com/submit/?U=<%=adres%>&T=<%=title%>">
								<img alt="Netscape" src="tema/v3_img/paylas/netscape.gif"border="0"></a>
								<a title="Netscape" target="_blank" href="http://www.netscape.com/submit/?U=<%=adres%>&T=<%=title%>">
								Netscape</a>
								<br>
								<a title="Netvouz" target="_blank" href="http://www.netvouz.com/action/submitBookmark?url=<%=adres%>&title=<%=title%>&description=TITLE">
								<img alt="Netvouz" src="tema/v3_img/paylas/netvouz.png"border="0"></a>
								<a title="Netvouz" target="_blank" href="http://www.netvouz.com/action/submitBookmark?url=<%=adres%>&title=<%=title%>&description=TITLE">
								Netvouz</a>
								<br>
								<a title="Oyyla" target="_blank" href="http://www.oyyla.com/gonder">
								<img alt="Oyyla" src="tema/v3_img/paylas/oyyla.gif"border="0"></a>
								<a title="Oyyla" target="_blank" href="http://www.oyyla.com/gonder">
								Oyyla</a>
								<br>
								<a title="RawSugar" target="_blank" href="http://www.rawsugar.com/tagger/?turl=<%=adres%>&tttl=<%=title%>">
								<img alt="RawSugar" src="tema/v3_img/paylas/rawsugar.png"border="0"></a>
								<a title="RawSugar" target="_blank" href="http://www.rawsugar.com/tagger/?turl=<%=adres%>&tttl=<%=title%>">
								RawSugar</a>
								<br>
								<a title="Reddit" target="_blank" href="http://reddit.com/submit?url=<%=adres%>&title=<%=title%>">
								<img alt="Reddit" src="tema/v3_img/paylas/reddit.png"border="0"></a>
								<a title="Reddit" target="_blank" href="http://reddit.com/submit?url=<%=adres%>&title=<%=title%>">
								Reddit</a>
								<br>
								<a title="Scuttle" target="_blank" href="http://www.scuttle.org/bookmarks.php/maxpower?action=add&address=<%=adres%>&title=<%=title%>&description=TITLE">
								<img alt="Scuttle" src="tema/v3_img/paylas/scuttle.png"border="0"></a>
								<a title="Scuttle" target="_blank" href="http://www.scuttle.org/bookmarks.php/maxpower?action=add&address=<%=adres%>&title=<%=title%>&description=TITLE">
								Scuttle</a>
								<br>
								<a title="Shadows" target="_blank" href="http://www.shadows.com/features/tcr.htm?url=<%=adres%>&title=<%=title%>">
								<img alt="Shadows" src="tema/v3_img/paylas/shadows.png"border="0"></a>
								<a title="Shadows" target="_blank" href="http://www.shadows.com/features/tcr.htm?url=<%=adres%>&title=<%=title%>">
								Shadows</a>
								<br>
								<a title="Simpy" target="_blank" href="http://www.simpy.com/simpy/LinkAdd.do?href=<%=adres%>&title=<%=title%>">
								<img alt="Simpy" src="tema/v3_img/paylas/simpy.png"border="0"></a>
								<a title="Simpy" target="_blank" href="http://www.simpy.com/simpy/LinkAdd.do?href=<%=adres%>&title=<%=title%>">
								Simpy</a>
							</ul>
									</td>
									<td width="33%" valign="top">
								<a title="Smarking" target="_blank" href="http://smarking.com/editbookmark/?url=<%=adres%>&description=<%=title%>">
								<img alt="Smarking" src="tema/v3_img/paylas/smarking.png"border="0"></a>
								<a title="Smarking" target="_blank" href="http://smarking.com/editbookmark/?url=<%=adres%>&description=<%=title%>">
								Smarking</a>
								<br>
								<a title="Spurl" target="_blank" href="http://www.spurl.net/spurl.php?url=<%=adres%>&title=<%=title%>">
								<img alt="Spurl" src="tema/v3_img/paylas/spurl.png"border="0"></a>
								<a title="Spurl" target="_blank" href="http://www.spurl.net/spurl.php?url=<%=adres%>&title=<%=title%>">
								Spurl</a>
								<br>
								<a title="Stumble Upon" target="_blank" href="http://www.stumbleupon.com/submit?url=<%=adres%>&title=<%=title%>">
								<img alt="Stumble Upon" src="tema/v3_img/paylas/su.png"border="0"></a>
								<a title="Stumble Upon" target="_blank" href="http://www.stumbleupon.com/submit?url=<%=adres%>&title=<%=title%>">
								Stumble Upon</a>
								<br>
								<a title="TailRank" target="_blank" href="http://tailrank.com/share/?text=&link_href=<%=adres%>&title=<%=title%>">
								<img alt="TailRank" src="tema/v3_img/paylas/tailrank.png"border="0"></a>
								<a title="TailRank" target="_blank" href="http://tailrank.com/share/?text=&link_href=<%=adres%>&title=<%=title%>">
								TailRank</a>
								<br>
								<a title="Technorati" target="_blank" href="http://www.technorati.com/faves?add=<%=adres%>">
								<img alt="Technorati" src="tema/v3_img/paylas/technorati.gif"border="0"></a>
								<a title="Technorati" target="_blank" href="http://www.technorati.com/faves?add=<%=adres%>">
								Technorati</a>
								<br>
								<a title="Tusul" target="_blank" href="http://www.tusul.com/submit.php">
								<img alt="Tusul" src="tema/v3_img/paylas/tusul.gif"border="0"></a>
								<a title="Tusul" target="_blank" href="http://www.tusul.com/submit.php">
								Tusul</a>
								<br>
								<a title="Wink" target="_blank" href="http://wink.com/_/tag?url=<%=adres%>&doctitle=<%=title%>">
								<img alt="Wink" src="tema/v3_img/paylas/wink.png"border="0"></a>
								<a title="Wink" target="_blank" href="http://wink.com/_/tag?url=<%=adres%>&doctitle=<%=title%>">
								Wink</a>
								<br>
								<a title="Wists" target="_blank" href="http://wists.com/r.php?c=&r=<%=adres%>&title=<%=title%>">
								<img alt="Wists" src="tema/v3_img/paylas/wists.png"border="0"></a>
								<a title="Wists" target="_blank" href="http://wists.com/r.php?c=&r=<%=adres%>&title=<%=title%>">
								Wists</a>
								<br>
								<a title="Yahoo! My Web" target="_blank" href="http://myweb2.search.yahoo.com/myresults/bookmarklet?u=<%=adres%>&=<%=title%>">
								<img alt="Yahoo! My Web" src="tema/v3_img/paylas/yahoomyweb.png"border="0"></a>
								<a title="Yahoo! My Web" target="_blank" href="http://myweb2.search.yahoo.com/myresults/bookmarklet?u=<%=adres%>&=<%=title%>">
								Yahoo! My Web</a>
								<br>
								<a title="Yumile!" target="_blank" href="http://www.yumiyum.org/imekle.php?iadres=<%=adres%>&ibaslik=<%=title%>">
								<img alt="Yumile!" src="tema/v3_img/paylas/yumiyum.gif"border="0"></a>
								<a title="Yumile!" target="_blank" href="http://www.yumiyum.org/imekle.php?iadres=<%=adres%>&ibaslik=<%=title%>">
								Yumile!</a>
								<br>
								<a title="Feed Me Links" target="_blank" href="http://feedmelinks.com/categorize?from=toolbar&op=submit&url=<%=adres%>&name=<%=title%>">
								<img alt="Feed Me Links" src="tema/v3_img/paylas/feedmelinks.png"border="0"></a>
								<a title="Feed Me Links" target="_blank" href="http://feedmelinks.com/categorize?from=toolbar&op=submit&url=<%=adres%>&name=<%=title%>">
								Feed Me Links</a>
								<br>
								<a title="Furl" target="_blank" href="http://www.furl.net/storeIt.jsp?u=<%=adres%>&t=<%=title%>">
								<img alt="Furl" src="tema/v3_img/paylas/furl.png"border="0"></a>
								<a title="Furl" target="_blank" href="http://www.furl.net/storeIt.jsp?u=<%=adres%>&t=<%=title%>">
								Furl</a>
								<br>
								<a title="Google Bookmarks" target="_blank" href="http://www.google.com/bookmarks/mark?op=edit&bkmk=<%=adres%>&title=<%=title%>">
								<img alt="Google Bookmarks" src="tema/v3_img/paylas/google_bmarks.gif"border="0"></a>
								<a title="Google Bookmarks" target="_blank" href="http://www.google.com/bookmarks/mark?op=edit&bkmk=<%=adres%>&title=<%=title%>">
								Google Bookmarks</a>
								</td>
								</tr>
<% End if %>
							</table>
					</td>
				</tr>
				<tr>
					<td colspan="3" width="550"><img src="tema/images/orta_alt.gif"></td>
				</tr>
			</table>