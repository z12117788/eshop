<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/conneshop.asp" -->
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
' *** Redirect if username exists
MM_flag="MM_insert"
If (CStr(Request(MM_flag)) <> "") Then
  MM_dupKeyRedirect="userisexist.asp"
  MM_rsKeyConnection=MM_conneshop_STRING
  MM_dupKeyUsernameValue = CStr(Request.Form("Email"))
  MM_dupKeySQL="SELECT Email FROM Customers WHERE Email='" & MM_dupKeyUsernameValue & "'"
  MM_adodbRecordset="ADODB.Recordset"
  set MM_rsKey=Server.CreateObject(MM_adodbRecordset)
  MM_rsKey.ActiveConnection=MM_rsKeyConnection
  MM_rsKey.Source=MM_dupKeySQL
  MM_rsKey.CursorType=0
  MM_rsKey.CursorLocation=2
  MM_rsKey.LockType=3
  MM_rsKey.Open
  If Not MM_rsKey.EOF Or Not MM_rsKey.BOF Then 
    ' the username was found - can not add the requested username
    MM_qsChar = "?"
    If (InStr(1,MM_dupKeyRedirect,"?") >= 1) Then MM_qsChar = "&"
    MM_dupKeyRedirect = MM_dupKeyRedirect & MM_qsChar & "requsername=" & MM_dupKeyUsernameValue
    Response.Redirect(MM_dupKeyRedirect)
  End If
  MM_rsKey.Close
End If
%>
<%
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) <> "") Then

  MM_editConnection = MM_conneshop_STRING
  MM_editTable = "Customers"
  MM_editRedirectUrl = "registerok.asp"
  MM_fieldsStr  = "Email|value|Pass|value|Pass_question|value|Pass_answer|value|Name|value|City|value|Address|value|Zip|value|Phone|value"
  MM_columnsStr = "Email|',none,''|Password|',none,''|PassQuestion|',none,''|PassAnswer|',none,''|Name|',none,''|City|',none,''|Address|',none,''|Zip|',none,''|Phone|',none,''"

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
' *** Insert Record: construct a sql insert statement and execute it

If (CStr(Request("MM_insert")) <> "") Then

  ' create the sql insert statement
  MM_tableValues = ""
  MM_dbValues = ""
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
      MM_tableValues = MM_tableValues & ","
      MM_dbValues = MM_dbValues & ","
    End if
    MM_tableValues = MM_tableValues & MM_columns(i)
    MM_dbValues = MM_dbValues & FormVal
  Next
  MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"

  If (Not MM_abortEdit) Then
    ' execute the insert
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
<html>
<head>
<title>网上商城</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="style.css" type="text/css">
<script language="JavaScript">
<!--
function isEmpty(s)
{
return ((s == null) || (s.length == 0))
}
function isWhitespace (s)
{
var whitespace = " \t\n\r";
var i;
// 以下代码判断是否有空字符
for (i = 0; i < s.length; i++)
{
var c = s.charAt(i);
if (whitespace.indexOf(c) >= 0)
{
return true;
}
}

return false;
}
function isCharsInBag (s, bag)
{
var i;
for (i = 0; i < s.length; i++)
{
var c = s.charAt(i);
if (bag.indexOf(c) == -1) return false;
}
return true;
}
function isEmail (s)
{
//判断Email是否为空
if (isEmpty(s))
{
window.alert("输入的E-mail地址不能为空，请输入！");
return false;
}
//判断Email中是否包含空格
if (isWhitespace(s))
{
window.alert("输入的E-mail地址中不能包含空格符，请重新输入！");
return false;
}
//判断Email地址长度
var i = 1;
var len = s.length;
if (len > 100)
{
window.alert("Email地址长度不能超过100位!");
return false;
}
pos1 = s.indexOf("@");
pos2 = s.indexOf(".");
pos3 = s.lastIndexOf("@");
pos4 = s.lastIndexOf(".");
//判断Email地址中是否包含符号 "@" 
if ((pos1 <= 0)||(pos1 == len)||(pos2 <= 0)||(pos2 == len))
{
window.alert("请输入有效的E-mail地址！");
return false;
}
else
{
if( (pos1 == pos2 - 1) || (pos1 == pos2 + 1)
|| ( pos1 != pos3 ) //find two @
|| ( pos4 < pos3 ) ) //. should behind the "@"
{
window.alert("请输入有效的E-mail地址！");
return false;
}
}
if ( !isCharsInBag( s, "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789.-_@"))
{
window.alert("email地址中只能包含字符ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789.-_@\n" + "请重新输入" );
return false;
}
//判断是否包含有效的字符
/*
var badChar = "><,[]{}?/+=|\\"\":;!#$%^&()`";
if ( isCharsInBag( s, badChar))
{
alert("请不要在email地址中输入字符 " + badChar + "\n" );
alert("请重新输入" );
return false;
}
*/
return true;
}
function checkdata() {
if (document.form1.Email.value=="") {
window.alert ("请输入您的E-mail地址 ！")
return false
}
if ( !isEmail(document.form1.Email.value) )
return false
if (document.form1.Pass.value=="") {
window.alert ("请输入您的密码 ！")
return false
}
if (document.form1.Pass.value.length<5) {
window.alert ("您的密码数必须大于4位 ！")
return false
}
if (document.form1.Pass.value.length>10) {
window.alert ("您的密码数必须小于10位 ！")
return false
}
if (document.form1.Comfpass.value=="") {
window.alert ("请输入您的确认密码 ！")
return false
}
if (document.form1.Pass.value!=document.form1. Comfpass.value) {
window.alert ("您的密码不一致 ！")
return false
}
if (document.form1.Pass_question.value=="") {
window.alert ("请输入您取回密码的问题 ！")
return false
}
if (document.form1. Pass_answer.value=="") {
window.alert ("请输入您取回密码的答案！")
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
          <td>　<a href="default.asp" class="red">首页</a> &gt; 新用户注册</td>
          <td>&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="760" border="0" cellspacing="1" cellpadding="0" align="center" bgcolor="#000000">
  <tr> 
    <td bgcolor="#FFFFFF"> 
      <form name="form1" method="POST" action="<%=MM_editAction%>" onSubmit="return checkdata()">
        　　　<img src="images/newregister.gif" width="190" height="30"> <br>
        <span class="productName">　　　　　如果您已经注册， 需要更改注册信息，<a href="cinfoupdate.asp">请点击这里</a></span> 
        <br>
        <br>
        <table width="80%" border="0" cellspacing="2" cellpadding="2" align="center" bgcolor="#CCCCCC">
          <tr bgcolor="#FFFFCC"> 
            <td colspan="2" height="24">　以下内容将作为你登录时的凭证，请认真填写！</td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="34%" align="right">电子邮件：</td>
            <td width="66%"> 
              <input type="text" name="Email" maxlength="100">
            </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="34%" align="right">密码：</td>
            <td width="66%"> 
              <input type="password" name="Pass" maxlength="10">
              密码需大于4位小于10位</td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="34%" align="right">确认密码：</td>
            <td width="66%"> 
              <input type="password" name="Comfpass" maxlength="10">
            </td>
          </tr>
          <tr bgcolor="#FFFFCC"> 
            <td colspan="2" height="24">　以下内容将用于你取回忘记的密码，请牢记～</td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="34%" align="right">取回密码问题：</td>
            <td width="66%"> 
              <input type="text" name="Pass_question" maxlength="100">
            </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="34%" align="right">取回密码答案：</td>
            <td width="66%"> 
              <input type="password" name="Pass_answer" maxlength="100">
            </td>
          </tr>
          <tr bgcolor="#FFFFCC"> 
            <td colspan="2" height="24">　请准确填写自己真实的相关信息，以便我们正确的为你提供服务。</td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="34%" align="right">真实姓名：</td>
            <td width="66%"> 
              <input type="text" name="Name" maxlength="50">
            </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="34%" align="right">城市：</td>
            <td width="66%"> 
              <input type="text" name="City" maxlength="50">
            </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="34%" align="right" height="17">详细地址：</td>
            <td height="17" width="66%"> 
              <textarea name="Address" rows="3" cols="30"></textarea>
            </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="34%" align="right">邮编：</td>
            <td width="66%"> 
              <input type="text" name="Zip" maxlength="10">
            </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="34%" align="right">电话：</td>
            <td width="66%"> 
              <input type="text" name="Phone" maxlength="50">
            </td>
          </tr>
          <tr bgcolor="#FFFFFF" valign="middle"> 
            <td colspan="2" align="center"> 
              <input type="submit" name="Submit" value="注册">
              　　　 
              <input type="reset" name="Submit2" value="重填">
            </td>
          </tr>
        </table>
        <p>&nbsp;</p>
        <input type="hidden" name="MM_insert" value="true">
      </form>
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
