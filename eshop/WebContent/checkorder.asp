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
Dim myorder__MMColParam
myorder__MMColParam = "1"
if (Session("MM_username")   <> "") then myorder__MMColParam = Session("MM_username")  
%>
<%
set myorder = Server.CreateObject("ADODB.Recordset")
myorder.ActiveConnection = MM_conneshop_STRING
myorder.Source = "SELECT Name,OrderID,OrderDate,Fulfilled  FROM Customers,Orders  WHERE Email = '" + Replace(myorder__MMColParam, "'", "''") + "' AND Customers.CustomerID=Orders.CustomerID  Order BY OrderDate Desc"
myorder.CursorType = 0
myorder.CursorLocation = 2
myorder.LockType = 3
myorder.Open()
myorder_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
myorder_numRows = myorder_numRows + Repeat1__numRows
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then MM_removeList = MM_removeList & "&" & MM_paramName & "="
MM_keepURL="":MM_keepForm="":MM_keepBoth="":MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each Item In Request.QueryString
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & NextItem & Server.URLencode(Request.QueryString(Item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each Item In Request.Form
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & NextItem & Server.URLencode(Request.Form(Item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
if (MM_keepBoth <> "") Then MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
if (MM_keepURL <> "")  Then MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
if (MM_keepForm <> "") Then MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
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
      <table width="80%" border="0" cellspacing="2" cellpadding="2" align="center">
        <tr align="center" valign="middle"> 
          <td height="24" class="productName" colspan="3"><%=(myorder.Fields.Item("Name").Value)%>，您好！以下是您的所有定单信息。</td>
        </tr>
        <tr align="center" valign="middle"> 
          <td height="24" class="productName" bgcolor="#5880A8"><font color="#FFFFFF">定单号</font></td>
          <td height="24" class="productName" bgcolor="#5880A8"><font color="#FFFFFF">下单日期</font></td>
          <td height="24" class="productName" bgcolor="#5880A8"><font color="#FFFFFF">是否处理</font></td>
        </tr>
        <% 
While ((Repeat1__numRows <> 0) AND (NOT myorder.EOF)) 
%>
        <tr align="center" valign="middle"> 
          <td height="24" bgcolor="#D9D9DB"><A HREF="orderdetail.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "OrderID=" & myorder.Fields.Item("OrderID").Value %>" class="productName"><%=(myorder.Fields.Item("OrderID").Value)%></A></td>
          <td height="24" bgcolor="#D9D9DB" class="productName"><%=(myorder.Fields.Item("OrderDate").Value)%></td>
          <td height="24" bgcolor="#D9D9DB"> 
            <% If myorder.Fields.Item("Fulfilled").Value = (-1) Then 'script %>
            <span class="productName">是</span> 
            <% End If ' end If myorder.Fields.Item("Fulfilled").Value = (-1) script %>
            <% If myorder.Fields.Item("Fulfilled").Value = (0) Then 'script %>
            <span class="productName">否</span> 
            <% End If ' end If myorder.Fields.Item("Fulfilled").Value = (0) script %>
          </td>
        </tr>
        <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  myorder.MoveNext()
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
myorder.Close()
%>
