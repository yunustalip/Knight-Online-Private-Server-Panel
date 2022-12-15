<%@LANGUAGE="VBSCRIPT" CODEPAGE="1254"%>
<!--#include virtual="/Connections/DSN_Egitim.asp" -->
<!--#include file="Connections/YASALEGITIM.asp" -->
<%
Dim dsn_rs
Dim dsn_rs_numRows

Set dsn_rs = Server.CreateObject("ADODB.Recordset")
dsn_rs.ActiveConnection = MM_DSN_Egitim_STRING
dsn_rs.Source = "SELECT * FROM setler"
dsn_rs.CursorType = 0
dsn_rs.CursorLocation = 2
dsn_rs.LockType = 1
dsn_rs.Open()

dsn_rs_numRows = 0
%>
<%
Dim dsn__less_rs
Dim dsn__less_rs_numRows

Set dsn__less_rs = Server.CreateObject("ADODB.Recordset")
dsn__less_rs.ActiveConnection = MM_YASALEGITIM_STRING
dsn__less_rs.Source = "SELECT * FROM setler"
dsn__less_rs.CursorType = 0
dsn__less_rs.CursorLocation = 2
dsn__less_rs.LockType = 1
dsn__less_rs.Open()

dsn__less_rs_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
dsn_rs_numRows = dsn_rs_numRows + Repeat1__numRows
%>
<%
Dim Repeat2__numRows
Dim Repeat2__index

Repeat2__numRows = -1
Repeat2__index = 0
dsn__less_rs_numRows = dsn__less_rs_numRows + Repeat2__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-9" />
<title>Untitled Document</title>
<style type="text/css">
<!--
.style1 {color: #000099}
-->
</style>
</head>

<body>
<table border="1">
  <tr>
    <td colspan="4" bgcolor="#FFFFCC"><div align="center"><em><strong>DREAMWEAVER 8 ÝLE VERÝTABANI ÇALIÞMASI </strong></em></div></td>
  </tr>
  <tr>
    <td bgcolor="#FFCCFF">No</td>
    <td bgcolor="#FFCCFF">kategori</td>
    <td bgcolor="#FFCCFF">set_adi</td>
    <td bgcolor="#FFCCFF">fiyat</td>
  </tr>
  <% While ((Repeat1__numRows <> 0) AND (NOT dsn_rs.EOF)) %>
    <tr>
      <td><%=(dsn_rs.Fields.Item("No").Value)%></td>
      <td><%=(dsn_rs.Fields.Item("kategori").Value)%></td>
      <td><%=(dsn_rs.Fields.Item("set_adi").Value)%></td>
      <td><%=(dsn_rs.Fields.Item("fiyat").Value)%></td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  dsn_rs.MoveNext()
Wend
%>
</table>
<p>&nbsp;</p>

<table border="1">
  <tr>
    <td bgcolor="#99FFCC"><strong>No</strong></td>
    <td bgcolor="#99FFCC"><strong>kategori</strong></td>
    <td bgcolor="#99FFCC"><strong>set_adi</strong></td>
    <td bgcolor="#99FFCC"><strong>fiyat</strong></td>
  </tr>
  <% While ((Repeat2__numRows <> 0) AND (NOT dsn__less_rs.EOF)) %>
    <tr>
      <td><span class="style1"><%=(dsn__less_rs.Fields.Item("No").Value)%></span></td>
      <td><span class="style1"><%=(dsn__less_rs.Fields.Item("kategori").Value)%></span></td>
      <td><span class="style1"><%=(dsn__less_rs.Fields.Item("set_adi").Value)%></span></td>
      <td><span class="style1"><%=(dsn__less_rs.Fields.Item("fiyat").Value)%></span></td>
    </tr>
    <% 
  Repeat2__index=Repeat2__index+1
  Repeat2__numRows=Repeat2__numRows-1
  dsn__less_rs.MoveNext()
Wend
%>
</table>
</body>
</html>
<%
dsn_rs.Close()
Set dsn_rs = Nothing
%>
<%
dsn__less_rs.Close()
Set dsn__less_rs = Nothing
%>
