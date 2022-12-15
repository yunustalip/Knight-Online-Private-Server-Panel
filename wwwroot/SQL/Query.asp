<%@ Language=VBScript %>
<%
	Response.Buffer=True
	Action=Request("Action")
	SqlQuery=Request("SqlQuery")
	DatabaseName=Request("dbname")	
	Srv=Request.Cookies("SQLADMIN")("Srv")
	Usr=Request.Cookies("SQLADMIN")("Usr")
	Pwd=Request.Cookies("SQLADMIN")("Pwd")
%>
<!-- #INCLUDE FILE="Library/WebGrid.asp" -->
<html>
	<head>
		<title>SQL Admin</title>
		<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-9">
		<meta http-equiv="Content-Language" content="tr">
		<meta name="Generator" content="">
		<meta name="Author" content="Sezer Turkmen">
		<meta name="Keywords" content="SQL Admin">
		<meta name="Description" content="SQL Admin">
		<meta name="Copyright" content="© Copyright 2005">
		<script language="JavaScript" src="Library/WebGrid.js"></script>
		<link href="Style/Style.css" type="text/css" rel="stylesheet">
		<link href="Style/WebGrid/Classic/Grid.css" type="text/css" rel="stylesheet">
	</head>
<body>
<table cellspacing="0" cellpadding="0" width="100%" height="100%" border="0" align="center">			
	<tr>				
		<td>					
			<table class="cssTableOutset" cellSpacing="3" cellPadding="0" width="100%" border="0">						
				<tr>							
					<td width="100%" style="font-size:12px;"><b>New Query</b></td>								
				</tr>					
			</table>				
		</td>			
	</tr>
	<tr><td height="1px"></td></tr>						
	<tr>				
		<td height="100%">
			<table class="cssTableOutset" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<tr>
					<td height="100%" vAlign="top" align="center">					
						<table width="100%" height="100%" cellSpacing="3" cellPadding="3" border="0">	
							<form name="frmSource" method="post" action="Query.asp?Action=Exec&dbname=<%=DatabaseName%>">
							<tr>
								<td width="100%" height="90%" align="center">
									<textarea wrap="off" name="SqlQuery" id="SqlQuery" cols="50" rows="50" style="width:100%;height:100%"><%=SqlQuery%></textarea>
								</td>
							</tr>							
							</form>
						</table>   					
					</td>						
				</tr>	
			</table>
		</td>						
	</tr>	
	<tr><td height="1px"></td></tr>						
	<tr>				
		<td height="30" valign="middle">
			<table class="cssTableOutset" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<tr>
					<td height="100%" vAlign="middle" align="left">					
						<table cellSpacing="1" cellPadding="0" border="0">
							<tr>
								<td vAlign="top"><button id="cmdOK" onclick="document.frmSource.submit();" style="width:75px;height:25px;">Execute</button></td>
							</tr>
						</table> 
					</td>						
				</tr>	
			</table>
		</td>						
	</tr>
<%
If Action="Exec" Then
	'On Error Resume Next 
	Set Conn=CreateObject("ADODB.Connection")
	Conn.Open "driver={SQL server};User Id=" & Usr & ";PASSWORD=" & Pwd & ";SERVER=" & Srv & ";UID=;APP=Microsoft Development Environment;DATABASE=" + DatabaseName
	Set Rs=Conn.Execute(SqlQuery)
	If Err<>0 Then 
		ErrMsg=Replace(Err.Description,"'","\'")
		Response.Write("<script>alert('" & ErrMsg & "');</script>")
	ElseIf Rs.State=1 Then 
%>
 	<tr>				
		<td height="300">
			<table height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<tr>
					<td height="100%" vAlign="top" align="center">						
						<table width="100%" height="100%" cellSpacing="0" cellPadding="0" border="0">
							<tr>
								<td width="100%" height="100%" align="center">
<%
		Response.Write(WebGrid("QueryResults", Rs, 50, "","",""))
		Rs.close		
%>
								</td>						
							</tr>					
						</table>
					</td>						
				</tr>					
			</table>
		</td>						
	</tr>	
<%
	Else 
		Response.Write("<script>alert('Your Query Successfuly Executed...')</script>") 
	End If 
	Conn.close 
End If 
%>
</table>
</body>
</html>

	 