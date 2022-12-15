<!--#include file="../_inc/conn.asp"-->
<!--#include file="../Function.asp"-->
<!--#include file="../md5.asp"-->
<%Response.CharSet = "windows-1254" 
Session.lcid = 1055
Session.CodePage = 1254

If Session("strAccountID")="" Or Not Session("durum")="esp" Then
Response.Redirect("login.asp")
Response.End
End If
	Response.Buffer=True
	ObjectType="Tables"
%>
<!-- #INCLUDE FILE="Library/WebGrid.asp" -->
<html>
	<head>
		<title>SQL Admin</title>
		<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-9">
		<meta http-equiv="Content-Language" content="tr">
		<meta GridName="Generator" content="">
		<meta GridName="Author" content="Sezer Turkmen">
		<meta GridName="Keywords" content="SQL Admin">
		<meta GridName="Description" content="SQL Admin">
		<script language="JavaScript" src="Library/WebGrid.js"></script>
		<link href="Style/Style.css" type="text/css" rel="stylesheet">
		<link href="Style/WebGrid/Classic/Grid.css" type="text/css" rel="stylesheet">
	</head>
	<style>	  		
		.active-column-0	{width: 300px;text-align:left}  
		.active-column-1	{width: 100px;text-align:center}
		.active-column-2	{width: 75px;text-align:center}
		.active-column-3	{width: 150px;text-align:left}
	</style>
	<script language="JavaScript">
		function jsWebGridDoubleClick(row)
		{
var chars=jsRandomString(); 
var dbname='<%=DatabaseName%>';
var objtype='<%=ObjectType%>';
var objid=row.getDataProperty("value",1);
var objname=row.getDataProperty("text",0);
var winLocation  = 'TableManager.asp?topen=table&tablename=' + objname;
var winFeatures = 'dialogWidth:640px; dialogHeight:480px; edge:none; center:yes; help:no; resizable:yes; scroll:no; status:no;';
var returnValue = window.open(winLocation,'detail','height='+(screen.availHeight*0.8)+',width='+(screen.availWidth*0.9)+',resizable=0,status=0,top=50,left=50,scrollbars=0,menubar=0,toolbar=0');
		}	  
		function jsRandomString() 
		{
var stringLength=50;
var randomString='';
var charList="0123456789ABCDEFGHIJKLMNOPQRSTUVWXTZ"; 
for (var i=0; i<stringLength; i++) {
	var randomNumber = Math.floor(Math.random() * charList.length);
	randomString += charList.substring(randomNumber,randomNumber+1);
}
return randomString;
		}
	</script>	
<body>
<table cellspacing="0" cellpadding="0" width="100%" height="100%" border="0" align="center">
	<tr>	
		<td>		
<table class="cssTableOutset" cellSpacing="3" cellPadding="0" width="100%" border="0">
	<tr>	
		<td width="100%" style="font-size:12px;"><b><%=ObjectType%></b></td>		
	</tr>		
</table>	
		</td>
	</tr>
	<tr>	
		<td height="100%">
<table height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
	<tr>
		<td height="100%" vAlign="top" align="center">
<table width="100%" height="100%" cellSpacing="0" cellPadding="0" border="0">
	<tr>
		<td width="100%" height="100%" align="center">	 
<%Select Case ObjectType
Case "Tables"
	sql = "SELECT sysobjects.name as 'NAME',sysobjects.id as ID, sysusers.name as 'OWNER', sysobjects.crdate as 'CREATE DATE' from sysobjects left outer join sysusers on sysobjects.uid=sysusers.uid WHERE sysobjects.xtype = 'U' ORDER BY sysobjects.name" 
	set rs=conne.execute(sql)
	Response.write(WebGrid("TablesGrid", rs, 50, "","jsWebGridDoubleClick",""))
	rs.close		
Case "Views"
	sql = "SELECT sysobjects.name as 'NAME',sysobjects.id as ID, sysusers.name as 'OWNER', sysobjects.crdate as 'CREATE DATE' from sysobjects left outer join sysusers on sysobjects.uid=sysusers.uid WHERE sysobjects.xtype = 'V' ORDER BY sysobjects.name" 
	set rs= conne.execute(sql)
	Response.write(WebGrid("ViewsGrid", rs, 50, "","jsWebGridDoubleClick",""))
	rs.close
Case "Procedures"
	sql = "SELECT sysobjects.name as 'NAME',sysobjects.id as ID, sysusers.name as 'OWNER', sysobjects.crdate as 'CREATE DATE' from sysobjects left outer join sysusers on sysobjects.uid=sysusers.uid WHERE sysobjects.xtype = 'P' ORDER BY sysobjects.name"
	set rs= conne.execute(sql)
	Response.write(WebGrid("ProceduresGrid", rs, 50, "","jsWebGridDoubleClick",""))
	rs.close
Case "Functions"
	sql = "SELECT sysobjects.name as 'NAME',sysobjects.id as ID, sysusers.name as 'OWNER', sysobjects.crdate as 'CREATE DATE' from sysobjects left outer join sysusers on sysobjects.uid=sysusers.uid WHERE sysobjects.xtype = 'FN' ORDER BY sysobjects.name"
	set rs= conne.execute(sql)
	Response.write(WebGrid("FunctionsGrid", rs, 50, "","jsWebGridDoubleClick",""))
	rs.close
End Select
conne.close
		%>
		</td>
	</tr>		
</table>
		</td>
	</tr>		
</table>
		</td>
	</tr>	
</table>
</body>
</html>