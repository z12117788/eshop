<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/conneshop.asp" -->
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers="admin,order"
MM_authFailedURL="logfail.asp"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
  If (false Or CStr(Session("MM_UserAuthorization"))="") Or _
         (InStr(1,MM_authorizedUsers,Session("MM_UserAuthorization"))>=1) Then
    MM_grantAccess = true
  End If
End If
If Not MM_grantAccess Then
  MM_qsChar = "?"
  If (InStr(1,MM_authFailedURL,"?") >= 1) Then MM_qsChar = "&"
  MM_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MM_referrer = MM_referrer & "?" & Request.QueryString()
  MM_authFailedURL = MM_authFailedURL & MM_qsChar & "accessdenied=" & Server.URLEncode(MM_referrer)
  Response.Redirect(MM_authFailedURL)
End If
%>
<%
Dim forderdetail__MMColParam
forderdetail__MMColParam = "1"
if (Request.QueryString("OrderID") <> "") then forderdetail__MMColParam = Request.QueryString("OrderID")
%>
<%
set forderdetail = Server.CreateObject("ADODB.Recordset")
forderdetail.ActiveConnection = MM_conneshop_STRING
forderdetail.Source = "SELECT * FROM OrderDetails WHERE OrderID = " + Replace(forderdetail__MMColParam, "'", "''") + ""
forderdetail.CursorType = 0
forderdetail.CursorLocation = 2
forderdetail.LockType = 3
forderdetail.Open()
forderdetail_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
forderdetail_numRows = forderdetail_numRows + Repeat1__numRows
%>
<html>
<head>
<title>网上商城管理系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
td {  font-family: "宋体"; font-size: 12px}
.mid {  font-family: "宋体"; font-size: 14px; font-weight: bold}
.fillbox {  font-family: "宋体"; font-size: 14px; font-weight: bold; color: #006600; border-color: black black #408C00; border-style: solid; border-top-width: 0px; border-right-width: 0px; border-bottom-width: 1px; border-left-width: 0px}
.sbutton {  font-family: "宋体"; font-size: 14px; color: #006600; background-color: #D0F0B3; font-weight: bold; padding-top: 1px; padding-bottom: 1px}
.bigfont {  font-family: "黑体"; font-size: 18px; font-weight: bold; color: #306898}
-->
</style>
</head>
<body bgcolor="#FFFFFF" text="#000000" topmargin="2">
<table width="760" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td><img src="images/admin_left.gif" width="382" height="98"></td>
    <td><img src="images/admin_right.gif" width="378" height="98" usemap="#Map" border="0"></td>
  </tr>
</table>
<table width="760" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td rowspan="3" bgcolor="#306898" width="1"></td>
    <td>　　　<img src="images/order.gif" width="129" height="27"></td>
    <td rowspan="3" bgcolor="#306898" width="1"></td>
  </tr>
  <tr> 
    <td align="center" valign="middle"> 
      <p class="bigfont">该定单详细内容</p>
      <p class="mid">定单编号：<%=(forderdetail.Fields.Item("OrderID").Value)%>　　总计费用：<%=(forderdetail.Fields.Item("TotalPrice").Value)%>元</p>
      <table width="95%" border="0" cellspacing="2" cellpadding="2" bgcolor="#7090B0">
        <tr align="center" valign="middle"> 
          <td class="mid" width="32%"><font color="#FFFFFF">商品编号</font></td>
          <td class="mid" width="46%"><font color="#FFFFFF">商品名称</font></td>
          <td class="mid" width="10%"><font color="#FFFFFF">数量</font></td>
          <td class="mid" width="12%"><font color="#FFFFFF">单价</font></td>
        </tr>
        <% 
While ((Repeat1__numRows <> 0) AND (NOT forderdetail.EOF)) 
%>
        <tr bgcolor="#FFFFFF"> 
          <td width="32%"><%=(forderdetail.Fields.Item("ProductID").Value)%></td>
          <td width="46%"><%=(forderdetail.Fields.Item("ProductName").Value)%></td>
          <td width="10%"><%=(forderdetail.Fields.Item("Quantity").Value)%></td>
          <td width="12%"><%=(forderdetail.Fields.Item("UnitPrice").Value)%>元</td>
        </tr>
        <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  forderdetail.MoveNext()
Wend
%>
        <tr> 
          <td height="3" colspan="4"></td>
        </tr>
      </table>
      <p>&nbsp;</p>
      <p>&nbsp;</p>
      <p>&nbsp;</p>
    </td>
  </tr>
  <tr> 
    <td height="4" bgcolor="#306898"></td>
  </tr>
</table>
<table width="760" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td background="images/admin_b1.gif" width="643" align="right" class="mid"><font color="#FFFFFF">copyright 
      2001 Powered By</font></td>
    <td width="117"><a href="http://www.fans8.com"><img src="images/admin_b2.gif" width="117" height="28" border="0" alt="..:::  狂迷俱乐部  :::.."></a></td>
  </tr>
</table>
<map name="Map"> 
  <area shape="rect" coords="134,20,234,41" href="cate_manage.asp">
  <area shape="rect" coords="254,20,355,43" href="order_manage.asp">
  <area shape="rect" coords="134,53,234,73" href="pro_manage.asp">
  <area shape="rect" coords="13,20,113,41" href="adminlogin.asp">
  <area shape="rect" coords="13,53,118,74" href="user_manage.asp">
  <area shape="rect" coords="254,52,355,74" href="ad_manage.asp">
</map>
</body>
</html>
<%
forderdetail.Close()
%>
