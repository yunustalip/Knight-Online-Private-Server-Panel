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





'Set the response buffer to true
Response.Buffer = True 



%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title>Batch Move Forum Topics</title>
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
<div align="center">
  <h1>Batch Move Forum Topics</h1><br />
    <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
    <br />
    From here you can move multiple Topics from one forum to another.<br />
    <br />
    Select the Forum you want Topics moved from and to by the date they where last posted in.<br />
  </p>
</div>
<form action="admin_batch_move_topics.asp<% = strQsSID1 %>" method="post" name="frmSelectForum" id="frmSelectForum" onsubmit="return confirm('Are you sure you want to move these topics?')">
  <table width="100%"  height="108" border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
    <tr>
      <td width="51%" valign="top" class="tableLedger">Move Topics from</td>
    </tr>
    <tr>
      <td class="tableRow"><strong>
        <select name="FFID">
          <option value="0" selected="selected">All Forums</option><%

'Read in the forum name from the database
'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT " & strDbTable & "Forum.Forum_name, " & strDbTable & "Forum.Forum_ID FROM " & strDbTable & "Category, " & strDbTable & "Forum WHERE " & strDbTable & "Category.Cat_ID = " & strDbTable & "Forum.Cat_ID ORDER BY " & strDbTable & "Category.Cat_order ASC, " & strDbTable & "Forum.Forum_Order ASC;"

'Query the database
rsCommon.Open strSQL, adoCon

'Loop through all the froum in the database
Do while NOT rsCommon.EOF 

	'Display a link in the link list to the forum
	Response.Write vbCrLf & "<option value=" & CLng(rsCommon("Forum_ID")) & " "
	Response.Write ">" & rsCommon("Forum_name") & "</option>"		
			
	'Move to the next record in the recordset
	rsCommon.MoveNext
Loop


%>
        </select>
      </strong></td>
    </tr>
    <tr>
      <td class="tableLedger">Move Topics To</td>
    </tr>
    <tr>
      <td class="tableRow"><select name="TFID"><%
      	
'Move to first record to loop through froums again
rsCommon.MoveFirst


'Loop through all the froum in the database
Do while NOT rsCommon.EOF 

	'Display a link in the link list to the forum
	Response.Write vbCrLf & "<option value=" & CLng(rsCommon("Forum_ID")) & " "
	Response.Write ">" & rsCommon("Forum_name") & "</option>"		
			
	'Move to the next record in the recordset
	rsCommon.MoveNext
Loop

'Reset server objects
rsCommon.Close
Call closeDatabase()
%>
        </select>
      </td>
    </tr>
    <tr>
      <td class="tableLedger">Move Topics that haven't been posted in for</td>
    </tr>
    <tr>
      <td class="tableRow"><select name="days">
          <option value="0" selected>Now</option>
          <option value="7">1 Week</option>
          <option value="14">2 Weeks</option>
          <option value="31">1 Month</option>
          <option value="62">2 Months</option>
          <option value="124">4 Months</option>
          <option value="182">6 Months</option>
          <option value="279">9 Months</option>
          <option value="365">1 Year</option>
          <option value="730">2 Years</option>
          <option value="1095">3 Years</option>
        </select>
      </td>
    </tr>
    <tr>
      <td class="tableLedger">Select which type of topics to move</td>
    </tr>
    <tr>
      <td class="tableRow">
        <select name="priority">
          <option value="4" selected="selected">All Topics</option>
          <option value="0">Normal Topics Only</option>
          <option value="1">Sticky Topics Only</option>
          <option value="2">Announcements Only</option>
        </select>
      </td>
    </tr>
    <tr>
      <td align="center" class="tableBottomRow">
      	<input type="hidden" name="formID" id="formID" value="<% = getSessionItem("KEY") %>" />
      	<input name="Submit" type="submit" value="Move Topics" /></td>
    </tr>
  </table>
</form>
<br />
<!-- #include file="includes/admin_footer_inc.asp" -->
