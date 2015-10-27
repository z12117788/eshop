<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="Connections/conneshop.asp" -->
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers=""
MM_authFailedURL="checkorder_login.asp"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
  If (true Or CStr(Session("MM_UserAuthorization"))="") Or _
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
Dim myorderdetail__MMColParam
myorderdetail__MMColParam = "1"
if (Request.QueryString("OrderID") <> "") then myorderdetail__MMColParam = Request.QueryString("OrderID")
%>
<%
set myorderdetail = Server.CreateObject("ADODB.Recordset")
myorderdetail.ActiveConnection = MM_conneshop_STRING
myorderdetail.Source = "SELECT * FROM OrderDetails WHERE OrderID = " + Replace(myorderdetail__MMColParam, "'", "''") + ""
myorderdetail.CursorType = 0
myorderdetail.CursorLocation = 2
myorderdetail.LockType = 3
myorderdetail.Open()
myorderdetail_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
myorderdetail_numRows = myorderdetail_numRows + Repeat1__numRows
%>
<html>
<head>
<title>网上商城</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="style.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" topmargin="2">
<table width="760" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td background="images/topback.gif" width="130"><img src="images/sitelogo.gif" height="88"></td>
    <td background="images/topback.gif" width="500" align="center" valign="middle"><a href="http://www.fans8.com" target="_blank"><img src="images/fans8.gif" width="468" height="60" border="0"></a> 
    </td>
    <td background="images/topback.gif" width="130"> <!-- #BeginLibraryItem "/Library/custm.lbi" --><table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td valign="middle" align="center"><a href="cart.asp"><img src="images/button_cart.gif" width="87" height="18" border="0"></a></td>
        </tr>
        <tr> 
          <td valign="middle" align="center"><a href="checkorder_login.asp"><img src="images/button_ddcx.gif" width="87" height="18" border="0"></a></td>
        </tr>
        <tr> 
          
    <td valign="middle" align="center"><a href="customer_register.asp"><img src="images/button_regist.gif" width="87" height="18" border="0"></a></td>
        </tr>
      </table><!-- #EndLibraryItem --></td>
  </tr>
</table>
<table width="760" border="0" cellspacing="1" cellpadding="0" align="center" bgcolor="#000000">
  <tr> 
    <td bgcolor="#FF9900" height="22" valign="middle" align="center"> <!-- #BeginLibraryItem "/Library/nav.lbi" --><table width="80%" border="0" cellspacing="2" cellpadding="2">
          <tr align="center" valign="middle"> 
            <td><a href="newproduct.asp" class="white">新品快递</a></td>
            
    <td><a href="commend.asp" class="white">重点推荐</a></td>
            
    <td><a href="bestsell.asp" class="white">销售排行</a></td>
            
    <td><a href="bestprice.asp" class="white">特价商品</a></td>
          </tr>
        </table><!-- #EndLibraryItem --></td>
  </tr>
  <tr> 
    <td bgcolor="#FFCC66" height="22"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td> 　<a href="default.asp" class="red">首页</a> &gt; 定单查询</td>
          <td>&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="760" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td> 　　<img src="images/checkorder.gif" width="190" height="30"> 
      <table width="90%" border="0" cellspacing="2" cellpadding="2" align="center">
        <tr align="center" valign="middle"> 
          <td height="24" class="productName" colspan="4">定单号：<%=(myorderdetail.Fields.Item("OrderID").Value)%>　　总计：<%=(myorderdetail.Fields.Item("TotalPrice").Value)%>元</td>
        </tr>
        <tr align="center" valign="middle"> 
          <td height="24" class="productName" bgcolor="#5880A8"><font color="#FFFFFF">商品编号</font></td>
          <td height="24" class="productName" bgcolor="#5880A8"><font color="#FFFFFF">商品名称</font></td>
          <td height="24" class="productName" bgcolor="#5880A8"><font color="#FFFFFF">单价</font></td>
          <td height="24" class="productName" bgcolor="#5880A8"><font color="#FFFFFF">购买数量</font></td>
        </tr>
        <% 
While ((Repeat1__numRows <> 0) AND (NOT myorderdetail.EOF)) 
%>
        <tr align="center" valign="middle"> 
          <td height="24" bgcolor="#D9D9DB"><%=(myorderdetail.Fields.Item("ProductID").Value)%></td>
          <td height="24" bgcolor="#D9D9DB"><%=(myorderdetail.Fields.Item("ProductName").Value)%></td>
          <td height="24" bgcolor="#D9D9DB"><%=(myorderdetail.Fields.Item("UnitPrice").Value)%>元</td>
          <td height="24" bgcolor="#D9D9DB"><%=(myorderdetail.Fields.Item("Quantity").Value)%></td>
        </tr>
        <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  myorderdetail.MoveNext()
Wend
%>
      </table>
      <p>&nbsp;</p>
    </td>
  </tr>
</table>
<br>
<!-- #BeginLibraryItem "/Library/bottm.lbi" --><table width="760" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td background="images/topback.gif" align="center" height="16"><font color="#FFFFFF">copyright 
      2001 Powered by Peter.HJ</font></td>
  </tr>
</table><!-- #EndLibraryItem --></body>
</html>
<%
myorderdetail.Close()
%>
