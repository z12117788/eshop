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
' *** Update Record: set variables

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_conneshop_STRING
  MM_editTable = "SubCategories"
  MM_editColumn = "SubCategoryID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "subcate.asp"
  MM_fieldsStr  = "subcatename|value|cateid|value|subcatedes|value|subcateimg|value|subcatecss|value"
  MM_columnsStr = "SubCategoryName|',none,''|CategoryID|none,none,NULL|SubCategoryDes|',none,''|SubCategoryImg|',none,''|SubCategoryStyle|',none,''"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(i+1) = CStr(Request.Form(MM_fields(i)))
  Next

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
' *** Update Record: construct a sql update statement and execute it

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql update statement
  MM_editQuery = "update " & MM_editTable & " set "
  For i = LBound(MM_fields) To UBound(MM_fields) Step 2
    FormVal = MM_fields(i+1)
    MM_typeArray = Split(MM_columns(i+1),",")
    Delim = MM_typeArray(0)
    If (Delim = "none") Then Delim = ""
    AltVal = MM_typeArray(1)
    If (AltVal = "none") Then AltVal = ""
    EmptyVal = MM_typeArray(2)
    If (EmptyVal = "none") Then EmptyVal = ""
    If (FormVal = "") Then
      FormVal = EmptyVal
    Else
      If (AltVal <> "") Then
        FormVal = AltVal
      ElseIf (Delim = "'") Then  ' escape quotes
        FormVal = "'" & Replace(FormVal,"'","''") & "'"
      Else
        FormVal = Delim + FormVal + Delim
      End If
    End If
    If (i <> LBound(MM_fields)) Then
      MM_editQuery = MM_editQuery & ","
    End If
    MM_editQuery = MM_editQuery & MM_columns(i) & " = " & FormVal
  Next
  MM_editQuery = MM_editQuery & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the update
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
Dim subcate__MMColParam
subcate__MMColParam = "1"
if (Request.QueryString("SubCategoryID") <> "") then subcate__MMColParam = Request.QueryString("SubCategoryID")
%>
<%
set subcate = Server.CreateObject("ADODB.Recordset")
subcate.ActiveConnection = MM_conneshop_STRING
subcate.Source = "SELECT * FROM SubCategories WHERE SubCategoryID = " + Replace(subcate__MMColParam, "'", "''") + ""
subcate.CursorType = 0
subcate.CursorLocation = 2
subcate.LockType = 3
subcate.Open()
subcate_numRows = 0
%>
<%
set cateid = Server.CreateObject("ADODB.Recordset")
cateid.ActiveConnection = MM_conneshop_STRING
cateid.Source = "SELECT CategoryID, CategoryName FROM Categories"
cateid.CursorType = 0
cateid.CursorLocation = 2
cateid.LockType = 3
cateid.Open()
cateid_numRows = 0
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
<script language="JavaScript">
<!--
function checkdata() {
if (document.form1.subcatename.value=="") {
window.alert ("请输入新增的子类名 ！")
return false
}
if (document.form1.subcatedes.value=="") {
window.alert ("请输入子类的描述 ！")
return false
}
if (document.form1.subcateimg.value=="") {
window.alert ("请输入子类的图片名 ！")
return false
}
if (document.form1.subcatecss.value=="") {
window.alert ("请输入子类的样式表名称 ！")
return false
}
return true
}
//-->
</script>
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
      <table width="80%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td align="center" valign="middle"> 
            <form name="form1" method="POST" action="<%=MM_editAction%>" onSubmit="return checkdata()">
              <table width="80%" border="0" cellspacing="2" cellpadding="2" bgcolor="#3868A0">
                <tr bgcolor="#FFFFFF"> 
                  <td width="27%" align="center" valign="middle" class="mid">子类名称</td>
                  <td width="73%"> 
                    <input type="text" name="subcatename" size="30" value="<%=(subcate.Fields.Item("SubCategoryName").Value)%>">
                  </td>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                  <td width="27%" align="center" valign="middle" class="mid">所属大类</td>
                  <td width="73%"> 
                    <select name="cateid">
                      <%
While (NOT cateid.EOF)
%>
                      <option value="<%=(cateid.Fields.Item("CategoryID").Value)%>"><%=(cateid.Fields.Item("CategoryName").Value)%></option>
                      <%
  cateid.MoveNext()
Wend
If (cateid.CursorType > 0) Then
  cateid.MoveFirst
Else
  cateid.Requery
End If
%>
                      <option value="<%=(subcate.Fields.Item("CategoryID").Value)%>" selected><%=(catename.Fields.Item("CategoryName").Value)%></option>
                    </select>
                  </td>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                  <td width="27%" align="center" valign="middle" class="mid">子类描述</td>
                  <td width="73%"> 
                    <textarea name="subcatedes" cols="50" rows="3"><%=(subcate.Fields.Item("SubCategoryDes").Value)%></textarea>
                  </td>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                  <td width="27%" align="center" valign="middle" class="mid">子类图片</td>
                  <td width="73%"> 
                    <input type="text" name="subcateimg" size="30" value="<%=(subcate.Fields.Item("SubCategoryImg").Value)%>">
                  </td>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                  <td width="27%" align="center" valign="middle" class="mid">子类样式</td>
                  <td width="73%"> 
                    <input type="text" name="subcatecss" size="30" value="<%=(subcate.Fields.Item("SubCategoryStyle").Value)%>">
                  </td>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                  <td colspan="2" align="center" valign="middle" class="mid"> 
                    <input type="image" border="0" name="imageField" src="images/update.gif" width="37" height="19">
                  </td>
                </tr>
              </table>
              <input type="hidden" name="MM_update" value="true">
              <input type="hidden" name="MM_recordId" value="<%= subcate.Fields.Item("SubCategoryID").Value %>">
            </form>
          </td>
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
subcate.Close()
%>
<%
cateid.Close()
%>
<%
catename.Close()
%>
