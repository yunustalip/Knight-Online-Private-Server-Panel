<%@ Language=VBScript %>
<%
	Response.Buffer=true 	
	Srv=Request.Cookies("SQLADMIN")("Srv")
	Usr=Request.Cookies("SQLADMIN")("Usr")
	Pwd=Request.Cookies("SQLADMIN")("Pwd")
%>
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
		<script language="JavaScript" src="Library/WebTree.js"></script>
		<link href="Style/WebTree/WebTree.css" type="text/css" rel="stylesheet">
		<link href="Style/Style.css" type="text/css" rel="stylesheet">
	</head>		
	<style>			
	HR {
		width:100%;
		color:#808080;
		height:2px;
		border:1px solid #808080;
		clear:both;	 
	}
	</style>	 
	<body leftmargin="5" topmargin="0">
		<table cellSpacing="1" cellPadding="1" width="100%" height="100%" border="0" >	
			<tr>
				<td width="100%" height="100%" vAlign="top">
					<table width="100%" height="100%" cellSpacing="0" cellPadding="0" border="0">
						<tr>
							<td width="100%" height="100%" vAlign="top">
							<script type="text/javascript" language="JavaScript">
							var webTree=new TreeView('webTree',false,true,true,false,'jsNodeClick','');
							webTree.add(0,-1,'<B><%=UCASE(Srv)%></B>&nbsp;(<%=Usr%>)<HR>');							
							<%
							childIndex=1
							set cnn=CreateObject("ADODB.Connection")
							sql="SELECT * FROM MASTER.DBO.SYSDATABASES"
							cnn.open "driver={SQL server};User Id=" & Usr & ";PASSWORD=" & Pwd & ";SERVER=" & Srv & ";UID=;APP=Microsoft Development Environment"
							set rs=cnn.execute(sql)
							while not rs.eof
								DbName=rs(0)
								Response.Write("webTree.add(" & childIndex+0 & ",0,'"  & DbName &  "','"  & DbName &  "');")
								Response.Write("webTree.add(" & childIndex+1 & "," & childIndex & ",'Tables','"  & DbName &  "');" )
								Response.Write("webTree.add(" & childIndex+2 & "," & childIndex & ",'Views','"  & DbName &  "');" )
								Response.Write("webTree.add(" & childIndex+3 & "," & childIndex & ",'Procedures','"  & DbName &  "');" )
								Response.Write("webTree.add(" & childIndex+4 & "," & childIndex & ",'Functions','"  & DbName &  "');" )
								Response.Write("webTree.add(" & childIndex+5 & "," & childIndex & ",'New Query','"  & DbName &  "');" )
								childIndex=childIndex+6 
								rs.movenext
							wend   
							rs.close
							set rs=Nothing	
							cnn.close  
							%>		
							document.write(webTree);
							</script>								
							</td>
						</tr>
					</table>
				</TD>
			</TR>
		</table>	
	</body>
</html>
<script language="JavaScript">
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
	function jsNodeClick(id,objType,dbName)
	{
		var chars=jsRandomString(); 
		if (objType=="New Query")
		{
			window.open('Query.asp?ID='+chars+'&dbname=' + dbName + '&objtype=' + objType+ '&objname=' + objType,'Main','');
		}
		else
		{
			window.open('List.asp?ID='+chars+'&dbname=' + dbName + '&objtype=' + objType,'Main','');
		}
	}	  	
</script>	 