<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="admin_common.asp" -->
<%
'****************************************************************************************
'**  Copyright Notice    
'**
'**  Web Wiz Forums(TM)
'**  http://www.webwizforums.com
'**                            
'**  Copyright (C)2001-2011 Web Wiz Ltd. All Rights Reserved.
'**  
'**  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS UNDER LICENSE FROM WEB WIZ LTD.
'**  
'**  IF YOU DO NOT AGREE TO THE LICENSE AGREEMENT THEN WEB WIZ LTD. IS UNWILLING TO LICENSE 
'**  THE SOFTWARE TO YOU, AND YOU SHOULD DESTROY ALL COPIES YOU HOLD OF 'WEB WIZ' SOFTWARE
'**  AND DERIVATIVE WORKS IMMEDIATELY.
'**  
'**  If you have not received a copy of the license with this work then a copy of the latest
'**  license contract can be found at:-
'**
'**  http://www.webwiz.co.uk/license
'**
'**  For more information about this software and for licensing information please contact
'**  'Web Wiz' at the address and website below:-
'**
'**  Web Wiz Ltd, Unit 10E, Dawkins Road Industrial Estate, Poole, Dorset, BH15 4JD, England
'**  http://www.webwiz.co.uk
'**
'**  Removal or modification of this copyright notice will violate the license contract.
'**
'****************************************************************************************



'*************************** SOFTWARE AND CODE MODIFICATIONS **************************** 
'**
'** MODIFICATION OF THE FREE EDITIONS OF THIS SOFTWARE IS A VIOLATION OF THE LICENSE  
'** AGREEMENT AND IS STRICTLY PROHIBITED
'**
'** If you wish to modify any part of this software a license must be purchased
'**
'****************************************************************************************



'Clean up
Call closeDatabase()

%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>Forum Kontrol Paneli Men�s�</title>
<meta name="generator" content="Web Wiz Forums" />
<%
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
Response.Write("<!--//" & _
vbCrLf & "/* *******************************************************" & _
vbCrLf & "Software: Web Wiz Forums(TM) ver. " & strVersion & "" & _
vbCrLf & "Info: http://www.webwizforums.com" & _
vbCrLf & "Copyright: (C)2001-2011 Web Wiz Ltd. All rights reserved" & _
vbCrLf & "******************************************************* */" & _
vbCrLf & "//-->" & vbCrLf & vbCrLf)
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******

%>
<!-- #include file="includes/admin_header_inc.asp" -->
   <table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
    <tr>
     <td class="tableLedger">Control Panel Menu </td>
    </tr>
    <tr>
     <td class="tableRow"><table width="100%" border="0" cellpadding="15" cellspacing="4">
       <tr>
        <td width="33%" align="center"><a href="admin_menu.asp?C=admin<% = strQsSID2 %>"><img src="forum_images/forum_setup.png" alt="Forum Ayar ve Y�netimi" border="0" /></a><br />
         <a href="admin_menu.asp?C=admin<% = strQsSID2 %>"><strong>Forum Ayar ve Y�netimi</strong></a></td>
        <td width="33%" align="center"><a href="admin_menu.asp?C=setup<% = strQsSID2 %>"><img src="forum_images/toolbox.png" alt="Ayar ve Konfig�rasyon" border="0" /></a> <br />
         <a href="admin_menu.asp?C=setup<% = strQsSID2 %>"><strong>Ayar ve Konfig�rasyon</strong></a> </td>
        <td width="33%" align="center"><a href="admin_menu.asp?C=members<% = strQsSID2 %>"><img src="forum_images/gorups_members.png" alt="Grup ve �ye Ara�lar�" border="0" /></a><br />
         <a href="admin_menu.asp?C=members<% = strQsSID2 %>"><strong>Grup ve �ye Ara�lar�</strong></a></td>
       </tr>
      </table>
       <table width="100%" border="0" cellpadding="15" cellspacing="4">
        <tr>
         <td width="33%" align="center"><a href="admin_menu.asp?C=security<% = strQsSID2 %>"><img src="forum_images/security.png" alt="G�venlik Ayarlar�" border="0" /></a><br />
          <a href="admin_menu.asp?C=security<% = strQsSID2 %>"><strong>G�venlik Ayarlar�</strong></a> </td>
         <td width="33%" align="center"><a href="admin_menu.asp?C=tools<% = strQsSID2 %>"><img src="forum_images/tools.png" alt="Forum Ara�lar�" border="0" /></a><br />
          <a href="admin_menu.asp?C=tools<% = strQsSID2 %>"><strong>Forum Ara�lar�</strong></a></td>
         <%
	If blnLCode  Then
	
%>
         <td width="33%" align="center"><a href="admin_menu.asp?C=upgrades<% = strQsSID2 %>""><img src="forum_images/webwizforums_box_sm.png" alt="Premium Edition Y�kseltme" border="0" /></a><br />
          <a href="admin_menu.asp?C=upgrades<% = strQsSID2 %>"><strong>Premium Edition Y�kseltme</strong></a></td>
         <%
        
	End If

%>
        </tr>
      </table></td>
    </tr>
   </table><%
   
'If the database is not moved tell the user there forum is not secure
If strDatabaseType = "Access" AND strDbPathAndName = Server.MapPath("database/wwForum.mdb") Then   
   
%>  
 <br />
 <table class="errorTable" cellspacing="1" cellpadding="3" align="center">
  <tr>
   <td><img src="<% = strImagePath %>error.png" alt="<% = strTxtError %>" /> <strong>G�venlik Zafiyeti Bulundu!!</strong>
    <br /><br />
    Access Veritaban� g�venli de�il
    <br /><br />
    <a href="http://www.webwizforums.comkb/" target="_blank">View information on how to secure your Forums's Access database.</a>
  </tr>
</table><%

End If


'If the database is not moved tell the user there forum is not secure
If LCase(strLoggedInUsername) = "administrator" AND lngLoggedInUserID = 1 AND blnDemoMode = False Then
   
 %>  
 <br />
 <table class="errorTable" cellspacing="1" cellpadding="3" align="center">
  <tr>
   <td><img src="<% = strImagePath %>error.png" alt="<% = strTxtError %>" /> <strong>G�venlik Zafiyeti Bulundu!!</strong>
    <br /><br />
    Your Admin Login Username and Password are not secure.
    <br /><br />
    <a href="admin_change_admin_username.asp">Update your Admin Login.</a>
  </tr>
</table><%  
	
End If
   


'If they want forum admin menu
If Request.QueryString("C") = "admin" Then

%>
   <br />
   <table width="100%" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
 <tr>
  <td class="tableLedger">Forum Y�netimi</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_view_forums.asp<% = strQsSID1 %>">Forum Y�netimi</a><br />
Yaratma, De�i�tirme, Forum ve forum kategorilerini silme, Forum detaylar�n� de�i�tirme, Forum izinlerini ayarlama, Kilitli forumlar, Forumlar� parola ile koruma, vb.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_statistics.asp<% = strQsSID1 %>">Forum �statistikleri</a><br />
Forum istatistiklerini g�sterir.</td>
 </tr>
</table>
   <%

End If

'If they want memebrs and group menu
If Request.QueryString("C") = "members" Then

%>
<br />
<table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
 <tr>
  <td class="tableLedger">Grup Y�netimi</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_view_groups.asp<% = strQsSID1 %>">Grup Y�netimi</a><br />
   Yarat, sil, detaylar�n� de�i�tir v.b.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_group_permissions_form.asp<% = strQsSID1 %>">Grup �zinleri Y�netimi</a><br />
   Burada forum d�zenleme, foruma giri�, mesaj yazma, anket yaratma v.b. izinleri d�zenleyebilirsiniz. </td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_group_points.asp<% = strQsSID1 %>">Set Group Points</a><br />
Setup or change The Group Point System for the number of Points Members get for various actions within the forums.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_view_ladder_groups.asp<% = strQsSID1 %>">Ladder Group Administration</a><br />
From this option you can create, delete or edit Ladder Groups.</td>
 </tr>
</table>
<br />
<table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
 <tr>
  <td class="tableLedger">�yelik Y�netimi</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_select_members.asp<% = strQsSID1 %>">�yelik Y�netimi</a><br />
   Burada forum �yelerini y�netebilirsiniz.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_register.asp<% = strQsSID1 %>">Yeni �ye Ekle </a><br />
   Burada yeni �ye ekleyebilirsiniz. </td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_find_user.asp<% = strQsSID1 %>">�ye �zinleri Y�netimi</a><br />
Burada forum d�zenleme, foruma giri�, mesaj yazma, anket yaratma v.b. izinleri d�zenleyebilirsiniz.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_suspend_registration.asp<% = strQsSID1 %>">�ye Kayd�n� Durdurma</a><br />
   Burada foruma yeni �ye kayd�n� durdurabilirsiniz.</td>
 </tr>
</table><%

End If

'If they want setup menu
If Request.QueryString("C") = "setup" Then

%>
<br />
<table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
 <tr>
  <td class="tableLedger">Konfig�rasyon Ara�lar� </td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_web_wiz_forums_settings.asp<% = strQsSID1 %>">Web Wiz Forums Genel Ayarlar</a><br />
   Configure general settings for your Forum.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_user_settings.asp<% = strQsSID1 %>">Kullan�c� Ayarlar� </a><br />
  Configure settings for your users.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_post_topic_configure.asp<% = strQsSID1 %>">Konu Ve Mesaj Ayarlar�</a><br />
  Configure the way Topics and Post look and feel.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_registration_settings.asp<% = strQsSID1 %>">Kay�t Ve Profil Ayarlar�</a><br />
  Configure what items are compulsory for registration, custom registration items, and how member profiles are display.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_pm_configure.asp<% = strQsSID1 %>">�zel Mesaj Ayarlar�</a><br />
  Configure the Private Messenger.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_calendar_configure.asp<% = strQsSID1 %>">Etkinlik Takvimi Ayarlar�</a><br />
  Configure the Events Calendar.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_chat_room_settings.asp<% = strQsSID1 %>">Sohbet Odas� Ayarlar�</a><br />
   Configure the settings for the Chat Room.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_mobile_settings.asp<% = strQsSID1 %>">Mobil Cihaz Ayarlar�</a><br />
   Configure SmartPhone and Tablet Mobile Device settings for your forum.</td>
 </tr>
  <tr class="tableRow">
  <td><a href="admin_seo_settings.asp<% = strQsSID1 %>">SEO (Search Engine Optimisation) Ayarlar�</a><br />
   Configure Search Engine Optimisations (SEO) to your forum and Analytics Tracking.</td>
 </tr>
  <tr class="tableRow">
  <td><a href="admin_rss_configure.asp<% = strQsSID1 %>">RSS Ayarlar�</a><br />
  RSS Feeds allow you to syndicate content from your forum. Use this option to configure how your RSS Feeds will work.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_newspad_configure.asp<% = strQsSID1 %>">NewsPad Ayarlar� (Toplu Email Ayarlar�)</a><br />
  Web Wiz NewsPad allows you to send eNewsletters and mass email your Forum Members.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_upload_configure.asp<% = strQsSID1 %>">Upload Ayarlar�</a><br />
Allow users to upload images and files in their posts, also setup Avatar uploading.</td>
</tr>
<tr class="tableRow">
  <td><a href="admin_email_notify_configure.asp<% = strQsSID1 %>">Email Ayarlar�</a><br />
Configure email settings and enable email features such as email notification, email account activation, etc.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_date_time_configure.asp<% = strQsSID1 %>">Tarih Ve Saat Ayarlar�</a><br />
This option allows you to set the global date and time format as well as any time off-set you need to have if your hosting is within a different time zone. </td>
 </tr>
<tr class="tableRow">
  <td><a href="admin_ads_settings.asp<% = strQsSID1 %>">Reklam Ayarlar�</a><br />
Monetize your forum content by affiliating links or add Text/Banners Ads to your forum including Google Adsense.</td>
 </tr>
</table>
<br />
<table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
 <tr>
  <td class="tableLedger">Forum Aray�z�</td>
 </tr>
 <tr>
  <td class="tableRow"><a href="admin_skin_configure.asp<% = strQsSID1 %>">Forum Aray�z� Ayarlar�</a><br />
   From this option you can apply a new skin to your forums and change the name, look, and feel of your forum. </td>
 </tr>
</table>
<%
End If

If (Request.QueryString("C") = "setup" OR Request.QueryString("C") = "admin") Then
	
%>
<br />
<%

If blnLCode Then

%>
 <tr>
  <td class="tableRow"><a href="http://www.webwizforums.com" target="_blank">About</a><br />
   About Web Wiz Web Wiz Forums.</td>
 </tr><%

End If

%>
</table>
<%

End If




'If they want security menu
If Request.QueryString("C") = "security" Then

%>
<br />
<table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
 <tr>
  <td class="tableLedger">�zinler</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_group_permissions_form.asp<% = strQsSID1 %>">Grup �zinlerini Ayarla</a><br />
Setup or change Group Permissions </td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_find_user.asp<% = strQsSID1 %>">�ye �zinlerini Ayarla</a><br />
From this option you can configure permissions  for individual Members.</td>
 </tr>
</table>
<br />
<table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
 <tr>
  <td class="tableLedger">Security and Anti-SPAM </td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_close_forums.asp<% = strQsSID1 %>">Forumlar� Kilitleme </a><br />
From this option you can Lock the Forums so that no-one can post, register, login. etc. on the forum. Useful if you forum comes under attack or you just need to close it for maintenance. </td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_spam_filter_configure.asp<% = strQsSID1 %>">SPAM Filtreleme</a><br />
   Configure the SPAM Filter with SPAM that you want to filter from your forum and the action to take if SPAM is detected.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_bad_word_filter_configure.asp<% = strQsSID1 %>">K�t� S�z Filtreleme �zellikleri</a><br /> 
   Configure the Swear Word Filter to block words that you feel are not appropriate to your forum.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_spam_configure.asp<% = strQsSID1 %>">Anti-SPAM Flood Kontrol �zellikleri</a><br />
Configure Anti-SPAM Flood Control settings so you don't get members spamming the forum with 1,000's of unwanted or abusive posts.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_ip_blocking.asp<% = strQsSID1 %>">IP Adresi Engelleme</a><br />
   Ban IP addresses and ranges. Anyone in a blacklisted ranges will only have limited capabilities within the forum system. </td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_email_domain_blocking.asp<% = strQsSID1 %>">E-posta Adresi Engelleme</a><br />
Use this option to block Email address and Email Domains from being registered on the board.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_suspend_registration.asp<% = strQsSID1 %>">�ye Kayd�n� Durdurma</a><br />
From this option you can Suspend New Users from Registering to use the forum.</td>
 </tr><%
 	
 	'If super admin let 'em change password
 	If lngLoggedInUserID <> 1 AND blnDemoMode = False Then
%>
 <tr class="tableRow">
  <td><a href="admin_change_admin_username.asp<% = strQsSID1 %>">Amin Kullan�c� Ad�n� Ve Parolas�n� De�i�tir</a><br />
This is highly recommended for higher security to prevent unauthorised persons access this Admin Control Panel.</td>
 </tr><%

	End If
%>
</table>
<%

End If

'If they want tools menu
If Request.QueryString("C") = "tools" Then

%>
<br />
<table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
 <tr>
  <td class="tableLedger">Forum Ara�lar�</td>
 </tr>
 <%
       
'If this is an access database show the compact and repair feature
If strDatabaseType = "Access" Then 

%>
 <tr class="tableRow">
  <td><a href="admin_compact_access_db.asp<% = strQsSID1 %>">Veritaban� D�zenleme ve Onarma</a><br />
   Form here you can compact and repair your Forums Access database to increase performance.</td>
 </tr>
 <%
  
End If

%>
 <tr class="tableRow">
  <td><a href="admin_resync_forum_post_count.asp<% = strQsSID1 %>">Konu ve Mesaj Saya�lar�n� G�ncelleme</a><br />
   Re-sync the Topic and Post Count display for the forum's</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_archive_topics_form.asp<% = strQsSID1 %>">Eski Konular� Kilitleme</a><br />
   Batch lock old Topics allows you to batch lock Topics that haven't been posted in for sometime.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_batch_delete_posts_form.asp<% = strQsSID1 %>">Konular� Silme</a><br />
   Clean out the Forum Database by batch deleting topics that have not been posted in for sometime.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_batch_move_topics_form.asp<% = strQsSID1 %>">Konular� Ta��ma</a><br />
   Batch move Topics from one forum to another.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_batch_delete_pm_form.asp<% = strQsSID1 %>">�zel Mesajlar� Silme</a><br />
   Clean out the Forum Database by batch deleting old Private Messages.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_batch_delete_members_form.asp<% = strQsSID1 %>">�yeleri Silme</a><br />
   Clean out the Forum Database by batch deleting Members who have never posted.</td>
 </tr>
</table>
<br />
<table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
 <tr>
  <td class="tableLedger">Aktarma Ara�lar�</td>
 </tr>
 <tr>
  <td class="tableRow"><a href="admin_db_import_members_form.asp<% = strQsSID1 %>">Ba�ka Bir Veritaban�ndan �ye Aktarma</a><br />
   This tool allows you to import members from an external database, such as another forum system, CMS, etc. </td>
 </tr>
</table><%

End If


If Request.QueryString("C") = "upgrades" Then
	
%>
<br />
<table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
 <tr>
  <td class="tableLedger">Upgrades</td>
 </tr><%
 	If blnLCode = True Then
%>
 <tr>
  <td class="tableRow"><a href="admin_license.asp<% = strQsSID1 %>">Premium Edition Upgrade</a><br />
   Upgrade Web Wiz Forums to the Full Premium Edition.</td>
 </tr><%
 
	End If
%>
 <tr>
  <td class="tableRow"><a href="admin_server_test.asp<% = strQsSID1 %>">Server Requirements Test </a><br />
Check that the server your site is hosted on and your web host have the correct requirements to run Web Wiz Forums. </td>
 </tr>
 <tr>
  <td class="tableRow"><a href="admin_update_check.asp<% = strQsSID1 %>">Check for updates</a><br />
Check for updates to the Forum Software</td>
 </tr>
 </tr>
</table>
<%

End If


%>
<!-- #include file="includes/admin_footer_inc.asp" -->

