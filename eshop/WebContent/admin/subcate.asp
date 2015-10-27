<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/conneshop.asp" -->
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers="admin,category"
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
set cate = Server.CreateObject("ADODB.Recordset")
cate.ActiveConnection = MM_conneshop_STRING
cate.Source = "SELECT CategoryID, CategoryName FROM Categories"
cate.CursorType = 0
cate.CursorLocation = 2
cate.LockType = 3
cate.Open()
cate_numRows = 0
%>
<%
set subcate = Server.CreateObject("ADODB.Recordset")
subcate.ActiveConnection = MM_conneshop_STRING
subcate.Source = "SELECT SubCategoryID, SubCategoryName, CategoryID FROM SubCategories"
subcate.CursorType = 0
subcate.CursorLocation = 2
subcate.LockType = 3
subcate.Open()
subcate_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
cate_numRows = cate_numRows + Repeat1__numRows
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
    <td>　　　<img src="images/category.gif" width="129" height="27"></td>
    <td rowspan="3" bgcolor="#306898" width="1"></td>
  </tr>
  <tr> 
    <td align="center" valign="middle"> 
      <p class="bigfont">商城子类管理</p>
      <% 
While ((Repeat1__numRows <> 0) AND (NOT cate.EOF)) 
%>
      <table width="60%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td><img src="images/square.gif" width="24" height="23"><span class="bigfont"><%=(cate.Fields.Item("CategoryName").Value)%></span></td>
        </tr>
        <%
FilterParam=cate.Fields.Item("CategoryID").Value
subcate.Filter="CategoryID="&FilterParam
While (NOT subcate.EOF)
%>
        <tr> 
          <td align="center" bgcolor="#BDEB9C"><span class="mid"><%=(subcate.Fields.Item("SubCategoryName").Value)%></span>　　<A HREF="subcate_update.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "SubCategoryID=" & subcate.Fields.Item("SubCategoryID").Value %>"><img src="images/update.gif" width="37" height="19" border="0"></A>　<A HREF="subcate_del.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "SubCategoryID=" & subcate.Fields.Item("SubCategoryID").Value %>"><img src="images/del.gif" width="37" height="19" border="0"></A>　<a href="subcate_add.asp"><img src="images/add.gif" width="37" height="19" border="0"></a></td>
        </tr>
        <%
subcate.MoveNext()
Wend
%>
      </table>
      <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  cate.MoveNext()
Wend
%>
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
cate.Close()
%>
<%
subcate.Close()
%>
