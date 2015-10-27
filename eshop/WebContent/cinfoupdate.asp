<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/conneshop.asp" -->
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers=""
MM_authFailedURL="cinfologin.asp"
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
  MM_editTable = "Customers"
  MM_editColumn = "CustomerID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "cinfoupok.asp"
  MM_fieldsStr  = "Name|value|Password|value|City|value|Address|value|Zip|value|Phone|value"
  MM_columnsStr = "Name|',none,''|Password|',none,''|City|',none,''|Address|',none,''|Zip|',none,''|Phone|',none,''"

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
Dim updateCinfo__MMColParam
updateCinfo__MMColParam = "1"
if (Session("MM_Username") <> "") then updateCinfo__MMColParam = Session("MM_Username")
%>
<%
set updateCinfo = Server.CreateObject("ADODB.Recordset")
updateCinfo.ActiveConnection = MM_conneshop_STRING
updateCinfo.Source = "SELECT * FROM Customers WHERE Email = '" + Replace(updateCinfo__MMColParam, "'", "''") + "'"
updateCinfo.CursorType = 0
updateCinfo.CursorLocation = 2
updateCinfo.LockType = 3
updateCinfo.Open()
updateCinfo_numRows = 0
%>
<html>
<head>
<title>网上商城</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="style.css" type="text/css">
<script language="JavaScript">
<!--
function checkdata() {
if (document.form1.Password.value=="") {
window.alert ("请输入您的密码 ！")
return false
}
if (document.form1.Password.value.length<5) {
window.alert ("您的密码数必须大于4位 ！")
return false
}
if (document.form1. Password.value.length>10) {
window.alert ("您的密码数必须小于10位 ！")
return false
}
if (document.form1.Name.value=="") {
window.alert ("请输入您的真实姓名 ！")
return false
}
if (document.form1.City.value=="") {
window.alert ("请输入所在城市 ！")
return false
}
if (document.form1.Address.value=="") {
window.alert ("请输入您的详细地址 ！")
return false
}
if (document.form1.Zip.value=="") {
window.alert ("请输入您的邮编 ！")
return false
}
if (document.form1.Phone.value=="") {
window.alert ("请输入您的电话 ！")
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
          <td>&nbsp;</td>
          <td>&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="760" border="0" cellspacing="1" cellpadding="0" align="center" bgcolor="#000000">
  <tr> 
    <td bgcolor="#FFFFFF"> 
      <p>　　　<img src="images/modifyinfo.gif" width="190" height="30"></p>
      <form method="post" action="<%=MM_editAction%>" name="form1" onSubmit="return checkdata()">
        <p align="center" class="productName">你的Email地址是<%=(updateCinfo.Fields.Item("Email").Value)%>，<br>
          请一定正确填写相关内容，以保证所购商品正确的配送。</p>
        <table align="center" bgcolor="#CCCCFF" width="390">
          <tr valign="baseline" bgcolor="#FFFFFF"> 
            <td nowrap align="right" width="50">姓名：</td>
            <td width="328"> 
              <input type="text" name="Name" value="<%=(updateCinfo.Fields.Item("Name").Value)%>" size="32">
            </td>
          </tr>
          <tr valign="baseline" bgcolor="#FFFFFF"> 
            <td nowrap align="right" width="50">密码：</td>
            <td width="328"> 
              <input type="password" name="Password" value="<%=(updateCinfo.Fields.Item("Password").Value)%>" size="32" maxlength="10">
            </td>
          </tr>
          <tr valign="baseline" bgcolor="#FFFFFF"> 
            <td nowrap align="right" width="50">城市：</td>
            <td width="328"> 
              <input type="text" name="City" value="<%=(updateCinfo.Fields.Item("City").Value)%>" size="32">
            </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td nowrap align="right" valign="top" width="50">地址：</td>
            <td valign="baseline" width="328"> 
              <textarea name="Address" cols="50" rows="5"><%=(updateCinfo.Fields.Item("Address").Value)%></textarea>
            </td>
          </tr>
          <tr valign="baseline" bgcolor="#FFFFFF"> 
            <td nowrap align="right" width="50">邮编：</td>
            <td width="328"> 
              <input type="text" name="Zip" value="<%=(updateCinfo.Fields.Item("Zip").Value)%>" size="32">
            </td>
          </tr>
          <tr valign="baseline" bgcolor="#FFFFFF"> 
            <td nowrap align="right" width="50">电话：</td>
            <td width="328"> 
              <input type="text" name="Phone" value="<%=(updateCinfo.Fields.Item("Phone").Value)%>" size="32">
            </td>
          </tr>
          <tr valign="baseline" bgcolor="#FFFFFF" align="center"> 
            <td nowrap colspan="2"> 
              <input type="submit" value="确认修改">
            </td>
          </tr>
        </table>
        <input type="hidden" name="MM_update" value="true">
        <input type="hidden" name="MM_recordId" value="<%= updateCinfo.Fields.Item("CustomerID").Value %>">
      </form>
      <p>&nbsp;</p>
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
updateCinfo.Close()
%>
