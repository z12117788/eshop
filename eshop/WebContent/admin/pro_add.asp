<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/conneshop.asp" -->
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers="admin,category,product"
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
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) <> "") Then

  MM_editConnection = MM_conneshop_STRING
  MM_editTable = "Products"
  MM_editRedirectUrl = "pro_manage.asp"
  MM_fieldsStr  = "pid|value|pname|value|cateid|value|subcateid|value|supply|value|author|value|description|value|price|value|listprice|value|hotdeal|value|pubdate|value|commend|value|simg|value|bimg|value|stock|value|sell|value|visit|value|weight|value"
  MM_columnsStr = "ProductID|',none,''|ProductName|',none,''|CategoryID|none,none,NULL|SubCategID|none,none,NULL|Supplier|',none,''|Author|',none,''|Description|',none,''|Price|none,none,NULL|ListPrice|none,none,NULL|HotDeal|none,none,NULL|PubDate|',none,''|Commend|none,none,NULL|sImgUrl|',none,''|bImgUrl|',none,''|Stock|none,none,NULL|Sell|none,none,NULL|Visits|none,none,NULL|UnitWeight|none,none,NULL"

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
set subcateid = Server.CreateObject("ADODB.Recordset")
subcateid.ActiveConnection = MM_conneshop_STRING
subcateid.Source = "SELECT SubCategoryID, SubCategoryName FROM SubCategories"
subcateid.CursorType = 0
subcateid.CursorLocation = 2
subcateid.LockType = 3
subcateid.Open()
subcateid_numRows = 0
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
if (document.form1.pid.value=="") {
window.alert ("请输入商品编号 ！")
return false
}
if (document.form1.pname.value=="") {
window.alert ("请输入商品名称 ！")
return false
}
if (document.form1.supply.value=="") {
window.alert ("请输入商品出版商 ！")
return false
}
if (document.form1.author.value=="") {
window.alert ("请输入商品作者/演员 ！")
return false
}
if (document.form1.description.value=="") {
window.alert ("请输入商品的描述 ！")
return false
}
if (document.form1.price.value=="") {
window.alert ("请输入商品原价 ！")
return false
}
if (document.form1.listprice.value=="") {
window.alert ("请输入商品现价 ！")
return false
}
if (document.form1.simg.value=="") {
window.alert ("请输入商品的小图片文件名 ！")
return false
}
if (document.form1.bimg.value=="") {
window.alert ("请输入商品的大图片文件名 ！")
return false
}
if (document.form1.stock.value=="") {
window.alert ("请输入商品库存量 ！")
return false
}
if (document.form1.sell.value=="") {
window.alert ("请输入商品销售量 ！")
return false
}
if (document.form1.visit.value=="") {
window.alert ("请输入商品访问量 ！")
return false
}
if (document.form1.weight.value=="") {
window.alert ("请输入商品重量 ！")
return false
}
return true
}
function setymd(){
var pub_y= document.form1.year.value
var pub_m= document.form1.month.value
var pub_d= document.form1.day.value
document.form1.pubdate.value = pub_y+"-"+pub_m+"-"+pub_d
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
    <td>　　　<img src="images/product.gif" width="129" height="27"></td>
    <td rowspan="3" bgcolor="#306898" width="1"></td>
  </tr>
  <tr> 
    <td align="center" valign="middle"> 
      <p class="bigfont">新增商品</p>
      <form name="form1" method="POST" action="<%=MM_editAction%>" onSubmit="return checkdata()">
        <table width="80%" border="0" cellspacing="2" cellpadding="2" bgcolor="#7090B8">
          <tr bgcolor="#FFFFFF"> 
            <td width="19%" align="center" valign="middle" height="18"><b>商品编号</b></td>
            <td> 
              <input type="text" name="pid" size="50">
            </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="19%" align="center" valign="middle" height="18"><b>商品名称</b></td>
            <td> 
              <input type="text" name="pname" size="50">
            </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="19%" align="center" valign="middle" height="18"><b>所属大类</b></td>
            <td> 
              <select name="cateid">
                <%
While (NOT cateid.EOF)
%>
                <option value="<%=(cateid.Fields.Item("CategoryID").Value)%>" ><%=(cateid.Fields.Item("CategoryName").Value)%></option>
                <%
  cateid.MoveNext()
Wend
If (cateid.CursorType > 0) Then
  cateid.MoveFirst
Else
  cateid.Requery
End If
%>
              </select>
            </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="19%" align="center" valign="middle" height="18"><b>所属子类</b></td>
            <td> 
              <select name="subcateid">
                <%
While (NOT subcateid.EOF)
%>
                <option value="<%=(subcateid.Fields.Item("SubCategoryID").Value)%>" ><%=(subcateid.Fields.Item("SubCategoryName").Value)%></option>
                <%
  subcateid.MoveNext()
Wend
If (subcateid.CursorType > 0) Then
  subcateid.MoveFirst
Else
  subcateid.Requery
End If
%>
              </select>
            </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="19%" align="center" valign="middle" height="18"><b>出版/发行商</b></td>
            <td> 
              <input type="text" name="supply" size="50">
            </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="19%" align="center" valign="middle" height="18"><b>作者/演员</b></td>
            <td> 
              <input type="text" name="author" size="50">
            </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="19%" align="center" valign="middle" height="18"><b>商品描述</b></td>
            <td> 
              <textarea name="description" cols="60" rows="6"></textarea>
            </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="19%" align="center" valign="middle" height="18"><b>原　价</b></td>
            <td> 
              <input type="text" name="price">
              元</td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="19%" align="center" valign="middle" height="18"><b>现　价</b></td>
            <td> 
              <input type="text" name="listprice">
              元</td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="19%" align="center" valign="middle" height="18"><b>是否打折</b></td>
            <td> 
              <input type="radio" name="hotdeal" value="-1">
              是　　　 
              <input type="radio" name="hotdeal" value="0" checked>
              否</td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="19%" align="center" valign="middle" height="18"><b>出版日期</b></td>
            <td> 
              <select name="year" onChange="setymd()">
                <option value="90">1990</option>
                <option value="91">1991</option>
                <option value="92">1992</option>
                <option value="93">1993</option>
                <option value="94">1994</option>
                <option value="95">1995</option>
                <option value="96">1996</option>
                <option value="97">1997</option>
                <option value="98">1998</option>
                <option value="99">1999</option>
                <option value="2000">2000</option>
                <option value="2001" selected>2001</option>
                <option value="2002">2002</option>
              </select>
              年 
              <select name="month" onChange="setymd()">
                <option value="1" selected>1</option>
                <option value="2">2</option>
                <option value="3">3</option>
                <option value="4">4</option>
                <option value="5">5</option>
                <option value="6">6</option>
                <option value="7">7</option>
                <option value="8">8</option>
                <option value="9">9</option>
                <option value="10">10</option>
                <option value="11">11</option>
                <option value="12">12</option>
              </select>
              月 
              <select name="day" onChange="setymd()">
                <option value="1" selected>1</option>
                <option value="2">2</option>
                <option value="3">3</option>
                <option value="4">4</option>
                <option value="5">5</option>
                <option value="6">6</option>
                <option value="7">7</option>
                <option value="8">8</option>
                <option value="9">9</option>
                <option value="10">10</option>
                <option value="11">11</option>
                <option value="12">12</option>
                <option value="13">13</option>
                <option value="14">14</option>
                <option value="15">15</option>
                <option value="16">16</option>
                <option value="18">17</option>
                <option value="19">19</option>
                <option value="20">20</option>
                <option value="21">21</option>
                <option value="22">22</option>
                <option value="23">23</option>
                <option value="24">24</option>
                <option value="25">25</option>
                <option value="26">26</option>
                <option value="27">27</option>
                <option value="28">28</option>
                <option value="29">29</option>
                <option value="30">30</option>
                <option value="31">31</option>
              </select>
              <input type="hidden" name="pubdate" value="01-1-1">
            </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="19%" align="center" valign="middle" height="18"><b>重点推荐</b></td>
            <td> 
              <input type="radio" name="commend" value="-1">
              是　　　 
              <input type="radio" name="commend" value="0" checked>
              否</td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="19%" align="center" valign="middle" height="18"><b>小图片名称</b></td>
            <td> 
              <input type="text" name="simg" size="50">
            </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="19%" align="center" valign="middle" height="18"><b>大图片名称</b></td>
            <td> 
              <input type="text" name="bimg" size="50">
            </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="19%" align="center" valign="middle" height="18"><b>库存量</b></td>
            <td> 
              <input type="text" name="stock">
            </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="19%" align="center" valign="middle" height="18"><b>销售量</b></td>
            <td> 
              <input type="text" name="sell">
            </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="19%" align="center" valign="middle" height="18"><b>访问量</b></td>
            <td> 
              <input type="text" name="visit">
            </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="19%" align="center" valign="middle" height="18"><b>重　量</b></td>
            <td> 
              <input type="text" name="weight">
              公斤</td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td colspan="2" align="center" valign="middle"> 
              <input type="image" border="0" name="imageField" src="images/add.gif" width="37" height="19">
            </td>
          </tr>
        </table>
        <input type="hidden" name="MM_insert" value="true">
      </form>
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
cateid.Close()
%>
<%
subcateid.Close()
%>
