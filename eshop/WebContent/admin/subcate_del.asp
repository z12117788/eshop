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
' *** Edit Operations: declare variables

MM_editAction = CStr(Request("URL"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Delete Record: declare variables

if (CStr(Request("MM_delete")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_conneshop_STRING
  MM_editTable = "SubCategories"
  MM_editColumn = "SubCategoryID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "subcate.asp"

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If
  
End If
%>
<%
' *** Delete Record: construct a sql delete statement and execute it

If (CStr(Request("MM_delete")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql delete statement
  MM_editQuery = "delete from " & MM_editTable & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the delete
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
Dim subtodel__MMColParam
subtodel__MMColParam = "1"
if (Request.QueryString("SubCategoryID") <> "") then subtodel__MMColParam = Request.QueryString("SubCategoryID")
%>
<%
set subtodel = Server.CreateObject("ADODB.Recordset")
subtodel.ActiveConnection = MM_conneshop_STRING
subtodel.Source = "SELECT * FROM SubCategories WHERE SubCategoryID = " + Replace(subtodel__MMColParam, "'", "''") + ""
subtodel.CursorType = 0
subtodel.CursorLocation = 2
subtodel.LockType = 3
subtodel.Open()
subtodel_numRows = 0
%>
<%
Dim catename__MMColParam
catename__MMColParam = "1"
if (Request.QueryString("SubCategoryID") <> "") then catename__MMColParam = Request.QueryString("SubCategoryID")
%>
<%
set catename = Server.CreateObject("ADODB.Recordset")
catename.ActiveConnection = MM_conneshop_STRING
catename.Source = "SELECT CategoryName  FROM SubCategories INNER JOIN Categories ON  SubCategories.CategoryID= Categories.CategoryID  WHERE SubCategoryID = " + Replace(catename__MMColParam, "'", "''") + ""
catename.CursorType = 0
catename.CursorLocation = 2
catename.LockType = 3
catename.Open()
catename_numRows = 0
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
      <p class="bigfont">你确定要删除此子类吗？</p>
      <form name="form1" method="POST" action="<%=MM_editAction%>">
        <table width="60%" border="0" cellspacing="2" cellpadding="2" bgcolor="#306898">
          <tr bgcolor="#FFFFFF"> 
            <td width="30%" class="mid" align="center" height="30">子类名称</td>
            <td width="70%" class="mid" height="30"><%=(subtodel.Fields.Item("SubCategoryName").Value)%></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="30%" class="mid" align="center" height="30">所属大类</td>
            <td width="70%" class="mid" height="30"><%=(catename.Fields.Item("CategoryName").Value)%></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="30%" class="mid" align="center" height="30">描　　述</td>
            <td width="70%" class="mid" height="30"><%=(subtodel.Fields.Item("SubCategoryDes").Value)%></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="30%" class="mid" align="center" height="30">图　　片</td>
            <td width="70%" class="mid" height="30"><%=(subtodel.Fields.Item("SubCategoryImg").Value)%></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="30%" class="mid" align="center" height="30">样　　式</td>
            <td width="70%" class="mid" height="30"><%=(subtodel.Fields.Item("SubCategoryStyle").Value)%></td>
          </tr>
          <tr bgcolor="#FFFFFF" align="center" valign="middle"> 
            <td colspan="2" height="30"> 
              <input type="image" border="0" name="imageField" src="images/ok.gif" width="37" height="19">
              　　　　<a href="subcate.asp"><img src="images/back.gif" width="37" height="19" border="0"></a></td>
          </tr>
        </table>
        <input type="hidden" name="MM_delete" value="true">
        <input type="hidden" name="MM_recordId" value="<%= subtodel.Fields.Item("SubCategoryID").Value %>">
      </form>
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
subtodel.Close()
%>
<%
catename.Close()
%>
