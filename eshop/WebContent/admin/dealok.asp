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

if(Session("MM_username") <> "") then updateorder__varfman = Session("MM_username")

if(Session("OrderID") <> "") then updateorder__varorderid = Session("OrderID")

%>
<%

set updateorder = Server.CreateObject("ADODB.Command")
updateorder.ActiveConnection = MM_conneshop_STRING
updateorder.CommandText = "UPDATE Orders  SET Fulfilled=Yes, FulfillTime=Now(), FulfillMan='" + Replace(updateorder__varfman, "'", "''") + "'  WHERE OrderID =" + Replace(updateorder__varorderid, "'", "''") + ""
updateorder.CommandType = 1
updateorder.CommandTimeout = 0
updateorder.Prepared = true
updateorder.Execute()

%>
<%
Dim cemail__MMColParam
cemail__MMColParam = "1"
if (Session("OrderID")  <> "") then cemail__MMColParam = Session("OrderID") 
%>
<%
set cemail = Server.CreateObject("ADODB.Recordset")
cemail.ActiveConnection = MM_conneshop_STRING
cemail.Source = "SELECT OrderID, Email  FROM Orders INNER JOIN Customers ON Orders.CustomerID=Customers.CustomerID  WHERE OrderID = " + Replace(cemail__MMColParam, "'", "''") + ""
cemail.CursorType = 0
cemail.CursorLocation = 2
cemail.LockType = 3
cemail.Open()
cemail_numRows = 0
%>
<%
If Not cemail.EOF Or Not cemail.BOF Then
Set JMail = Server.CreateObject("JMail.SMTPMail") 
JMail.ServerAddress = "mail.yourserver.com:25"
JMail.Sender = "you@yourserver.com" 
JMail.Subject = "您在Fans8商城的定单已经处理"
JMail.AddRecipient (cemail.Fields.Item("Email").Value)
JMail.Body = "尊敬的顾客：." & vbCrLf & vbCrLf
JMail.Body = JMail.Body & "您在Fans8商城的定单，定单号为" &(cemail.Fields.Item("OrderID").Value) &"已经处理。"
JMail.Body = JMail.Body & "请留意查收。"
JMail.Body = JMail.Body & "欢迎光临Fans8网上商城。"
JMail.Priority = 3
JMail.AddHeader "Originating-IP", Request.ServerVariables("REMOTE_ADDR")
JMail.Execute
End If
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
      <p>&nbsp;</p>
      <p class="bigfont">定单已经确认处理，并已经向下单顾客发送确认邮件。</p>
      <p class="bigfont"><a href="orderlist.asp">返回</a></p>
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
cemail.Close()
%>
<%
cemail.Close()
Session("OrderID")=""
%>

