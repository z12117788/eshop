<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/conneshop.asp" -->
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers="admin,advertise"
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
set showid = Server.CreateObject("ADODB.Recordset")
showid.ActiveConnection = MM_conneshop_STRING
showid.Source = "SELECT Max(AdID) AS Max_id,Min(AdID) AS min_id   FROM Advertise"
showid.CursorType = 0
showid.CursorLocation = 2
showid.LockType = 3
showid.Open()
showid_numRows = 0
%>
<%
min=showid.Fields.Item("Min_id").Value
max=showid.Fields.Item("Max_id").Value
dim idtoshow
Randomize Timer
idtoshow=Int((max-min+1)*Rnd+min)
set adtoshow = Server.CreateObject("ADODB.Recordset")
adtoshow.ActiveConnection = MM_conneshop_STRING
adtoshow.Source = "SELECT * FROM Advertise WHERE AdID="&idtoshow
adtoshow.CursorType = 0
adtoshow.CursorLocation = 2
adtoshow.LockType = 3
adtoshow.Open()
adtoshow_numRows = 0
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
    <td>　　　<img src="images/adver.gif" width="129" height="27"></td>
    <td rowspan="3" bgcolor="#306898" width="1"></td>
  </tr>
  <tr> 
    <td align="center" valign="middle"> 
      <p class="bigfont">当前网站广告显示</p>
      <table width="80%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <% If Not adtoshow.EOF Or Not adtoshow.BOF Then %>
          <td align="center" valign="middle"><a href="<%=(adtoshow.Fields.Item("AdURL").Value)%>" target="_blank"><img src="../adbanner/<%=(adtoshow.Fields.Item("PicName").Value)%>" width="468" height="60" border="0" alt="<%=(adtoshow.Fields.Item("AltInfo").Value)%>"></a></td>
          <% End If ' end Not adtoshow.EOF Or NOT adtoshow.BOF %>
        </tr>
        <tr> 
          <% If adtoshow.EOF And adtoshow.BOF Then %>
          <td align="center" valign="middle"><a href="http://www.fans8.com" target="_blank"><img src="../adbanner/fans8.gif" width="468" height="60" border="0" alt=".:: 狂迷俱乐部 ::."></a></td>
          <% End If ' end adtoshow.EOF And adtoshow.BOF %>
        </tr>
      </table>
      <p>刷新页面可以随机看到其他广告</p>
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
showid.Close()
%>
<%
adtoshow.Close()
%>
