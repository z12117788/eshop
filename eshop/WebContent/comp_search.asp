<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/conneshop.asp" -->
<%
Dim qsearch__VarSubCateID
qsearch__VarSubCateID = "1"
if (Request("h_subcateid")  <> "") then qsearch__VarSubCateID = Request("h_subcateid") 
%>
<%
Dim qsearch__VarPname
qsearch__VarPname = "%"
if (Request("tex_name") <> "") then qsearch__VarPname = Request("tex_name")
%>
<%
Dim qsearch__VarAuthor
qsearch__VarAuthor = "%"
if (Request("tex_author") <> "") then qsearch__VarAuthor = Request("tex_author")
%>
<%
Dim qsearch__VarSupply
qsearch__VarSupply = "%"
if (Request("tex_supply") <> "") then qsearch__VarSupply = Request("tex_supply")
%>
<%
Dim qsearch__VarPubDate
qsearch__VarPubDate = "1"
if (Request("menu_pub") <> "") then qsearch__VarPubDate = Request("menu_pub")
%>
<%
Dim qsearch__VarHotDeal
qsearch__VarHotDeal = "1"
if (Request("menu_hotdeal")  <> "") then qsearch__VarHotDeal = Request("menu_hotdeal") 
%>
<%
Dim qsearch__VarListPrice
qsearch__VarListPrice = "1"
if (Request("menu_price") <> "") then qsearch__VarListPrice = Request("menu_price")
%>
<%
set qsearch = Server.CreateObject("ADODB.Recordset")
qsearch.ActiveConnection = MM_conneshop_STRING
qsearch.Source = "SELECT ProductID, ProductName, Supplier, Author, sImgUrl  FROM Products  WHERE SubCategID=" + Replace(qsearch__VarSubCateID, "'", "''") + " AND ProductName LIKE '%" + Replace(qsearch__VarPname, "'", "''") + "%' AND Author LIKE '%" + Replace(qsearch__VarAuthor, "'", "''") + "%' AND Supplier LIKE '%" + Replace(qsearch__VarSupply, "'", "''") + "%'AND " + Replace(qsearch__VarPubDate, "'", "''") + " AND " + Replace(qsearch__VarHotDeal, "'", "''") + " AND " + Replace(qsearch__VarListPrice, "'", "''") + "  ORDER BY AddDate DESC"
qsearch.CursorType = 0
qsearch.CursorLocation = 2
qsearch.LockType = 3
qsearch.Open()
qsearch_numRows = 0
%>
<%
Dim HLooper1__numRows
HLooper1__numRows = 10
Dim HLooper1__index
HLooper1__index = 0
qsearch_numRows = qsearch_numRows + HLooper1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

' set the record count
qsearch_total = qsearch.RecordCount

' set the number of rows displayed on this page
If (qsearch_numRows < 0) Then
  qsearch_numRows = qsearch_total
Elseif (qsearch_numRows = 0) Then
  qsearch_numRows = 1
End If

' set the first and last displayed record
qsearch_first = 1
qsearch_last  = qsearch_first + qsearch_numRows - 1

' if we have the correct record count, check the other stats
If (qsearch_total <> -1) Then
  If (qsearch_first > qsearch_total) Then qsearch_first = qsearch_total
  If (qsearch_last > qsearch_total) Then qsearch_last = qsearch_total
  If (qsearch_numRows > qsearch_total) Then qsearch_numRows = qsearch_total
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (qsearch_total = -1) Then

  ' count the total records by iterating through the recordset
  qsearch_total=0
  While (Not qsearch.EOF)
    qsearch_total = qsearch_total + 1
    qsearch.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (qsearch.CursorType > 0) Then
    qsearch.MoveFirst
  Else
    qsearch.Requery
  End If

  ' set the number of rows displayed on this page
  If (qsearch_numRows < 0 Or qsearch_numRows > qsearch_total) Then
    qsearch_numRows = qsearch_total
  End If

  ' set the first and last displayed record
  qsearch_first = 1
  qsearch_last = qsearch_first + qsearch_numRows - 1
  If (qsearch_first > qsearch_total) Then qsearch_first = qsearch_total
  If (qsearch_last > qsearch_total) Then qsearch_last = qsearch_total

End If
%>
<%
' *** Move To Record and Go To Record: declare variables

Set MM_rs    = qsearch
MM_rsCount   = qsearch_total
MM_size      = qsearch_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  r = Request.QueryString("index")
  If r = "" Then r = Request.QueryString("offset")
  If r <> "" Then MM_offset = Int(r)

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  i = 0
  While ((Not MM_rs.EOF) And (i < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    i = i + 1
  Wend
  If (MM_rs.EOF) Then MM_offset = i  ' set MM_offset to the last possible record

End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  i = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or i < MM_offset + MM_size))
    MM_rs.MoveNext
    i = i + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = i
    If (MM_size < 0 Or MM_size > MM_rsCount) Then MM_size = MM_rsCount
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  ' move the cursor to the selected record
  i = 0
  While (Not MM_rs.EOF And i < MM_offset)
    MM_rs.MoveNext
    i = i + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
qsearch_first = MM_offset + 1
qsearch_last  = MM_offset + MM_size
If (MM_rsCount <> -1) Then
  If (qsearch_first > MM_rsCount) Then qsearch_first = MM_rsCount
  If (qsearch_last > MM_rsCount) Then qsearch_last = MM_rsCount
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
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
<%
' *** Move To Record: set the strings for the first, last, next, and previous links

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 0) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    params = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For i = 0 To UBound(params)
      nextItem = Left(params(i), InStr(params(i),"=") - 1)
      If (StrComp(nextItem,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & params(i)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then MM_keepMove = MM_keepMove & "&"
urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="
MM_moveFirst = urlStr & "0"
MM_moveLast  = urlStr & "-1"
MM_moveNext  = urlStr & Cstr(MM_offset + MM_size)
prev = MM_offset - MM_size
If (prev < 0) Then prev = 0
MM_movePrev  = urlStr & Cstr(prev)
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
<form name="form2" method="post" action="">
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
            <td>　<a href="default.asp" class="red">首页</a> &gt; 搜索商品</td>
            <td>&nbsp;</td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</form>
<table width="760" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td height="30" valign="middle"> 
      <% If Not qsearch.EOF Or Not qsearch.BOF Then %>
      &nbsp; 共搜索到<%=(qsearch_total)%>个相关商品，当前为第<%=(qsearch_first)%> 到<%=(qsearch_last)%>个商品，你可以使用<a href="advanced_search.asp" class="mark">高级组合搜索</a>进行更详细的查询 
      <% End If ' end Not qsearch.EOF Or NOT qsearch.BOF %>
      <br>
      <% If qsearch.EOF And qsearch.BOF Then %>
      对不起，没有搜索到你要查找的内容，你可以使用<a href="advanced_search.asp" class="mark">高级组合搜索</a>进行更详细的查询 
      <% End If ' end qsearch.EOF And qsearch.BOF %>
    </td>
  </tr>
  <tr> 
    <td> 
      <% If Not qsearch.EOF Or Not qsearch.BOF Then %>
      <table>
        <%
startrw = 0
endrw = HLooper1__index
numberColumns = 2
numrows = 5
while((numrows <> 0) AND (Not qsearch.EOF))
	startrw = endrw + 1
	endrw = endrw + numberColumns
 %>
        <tr align="center" valign="top"> 
          <%
While ((startrw <= endrw) AND (Not qsearch.EOF))
%>
          <td> 
            <table width="380" border="0" cellspacing="2" cellpadding="2">
              <tr> 
                <td rowspan="3" width="111" align="left" valign="top"><A HREF="product.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "ProductID=" & qsearch.Fields.Item("ProductID").Value %>"><img src="images/product/<%=(qsearch.Fields.Item("sImgUrl").Value)%>" border="0"></A></td>
                <td width="255" valign="middle" class="productName"><A HREF="product.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "ProductID=" & qsearch.Fields.Item("ProductID").Value %>"><%=(qsearch.Fields.Item("ProductName").Value)%></A></td>
              </tr>
              <tr> 
                <td width="255" valign="middle"><%=(qsearch.Fields.Item("Author").Value)%></td>
              </tr>
              <tr> 
                <td width="255" valign="middle"><%=(qsearch.Fields.Item("Supplier").Value)%></td>
              </tr>
              <tr bgcolor="#CCCCCC"> 
                <td height="1" colspan="2" align="left" valign="top"></td>
              </tr>
            </table>
          </td>
          <%
	startrw = startrw + 1
	qsearch.MoveNext()
	Wend
	%>
        </tr>
        <%
 numrows=numrows-1
 Wend
 %>
      </table>
      <% End If ' end Not qsearch.EOF Or NOT qsearch.BOF %>
    </td>
  </tr>
  <tr> 
    <td>&nbsp; 
      <table border="0" width="50%" align="center">
        <tr> 
          <td width="23%" align="center"> 
            <% If MM_offset <> 0 Then %>
            <a href="<%=MM_moveFirst%>" class="navi">第一页</a> 
            <% End If ' end MM_offset <> 0 %>
          </td>
          <td width="31%" align="center"> 
            <% If MM_offset <> 0 Then %>
            <a href="<%=MM_movePrev%>" class="navi">前一页</a> 
            <% End If ' end MM_offset <> 0 %>
          </td>
          <td width="23%" align="center"> 
            <% If Not MM_atTotal Then %>
            <a href="<%=MM_moveNext%>" class="navi">下一页</a> 
            <% End If ' end Not MM_atTotal %>
          </td>
          <td width="23%" align="center"> 
            <% If Not MM_atTotal Then %>
            <a href="<%=MM_moveLast%>" class="navi">最末页</a> 
            <% End If ' end Not MM_atTotal %>
          </td>
        </tr>
      </table>
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
qsearch.Close()
%>
