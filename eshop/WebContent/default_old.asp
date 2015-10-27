<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/conneshop.asp" -->
<%
Dim booksub__MMColParam
booksub__MMColParam = "1"
if (Request("MM_EmptyValue") <> "") then booksub__MMColParam = Request("MM_EmptyValue")
%>
<%
set booksub = Server.CreateObject("ADODB.Recordset")
booksub.ActiveConnection = MM_conneshop_STRING
booksub.Source = "SELECT SubCategoryID, SubCategoryName FROM SubCategories WHERE CategoryID = " + Replace(booksub__MMColParam, "'", "''") + ""
booksub.CursorType = 0
booksub.CursorLocation = 2
booksub.LockType = 3
booksub.Open()
booksub_numRows = 0
%>
<%
Dim videosub__MMColParam
videosub__MMColParam = "2"
if (Request("MM_EmptyValue") <> "") then videosub__MMColParam = Request("MM_EmptyValue")
%>
<%
set videosub = Server.CreateObject("ADODB.Recordset")
videosub.ActiveConnection = MM_conneshop_STRING
videosub.Source = "SELECT SubCategoryID, SubCategoryName FROM SubCategories WHERE CategoryID = " + Replace(videosub__MMColParam, "'", "''") + ""
videosub.CursorType = 0
videosub.CursorLocation = 2
videosub.LockType = 3
videosub.Open()
videosub_numRows = 0
%>
<%
Dim musicsub__MMColParam
musicsub__MMColParam = "3"
if (Request("MM_EmptyValue") <> "") then musicsub__MMColParam = Request("MM_EmptyValue")
%>
<%
set musicsub = Server.CreateObject("ADODB.Recordset")
musicsub.ActiveConnection = MM_conneshop_STRING
musicsub.Source = "SELECT SubCategoryID, SubCategoryName FROM SubCategories WHERE CategoryID = " + Replace(musicsub__MMColParam, "'", "''") + ""
musicsub.CursorType = 0
musicsub.CursorLocation = 2
musicsub.LockType = 3
musicsub.Open()
musicsub_numRows = 0
%>
<%
set HotProduct = Server.CreateObject("ADODB.Recordset")
HotProduct.ActiveConnection = MM_conneshop_STRING
HotProduct.Source = "SELECT ProductID, ProductName FROM Products ORDER BY Visits DESC"
HotProduct.CursorType = 0
HotProduct.CursorLocation = 2
HotProduct.LockType = 3
HotProduct.Open()
HotProduct_numRows = 0
%>
<%
set BestSell = Server.CreateObject("ADODB.Recordset")
BestSell.ActiveConnection = MM_conneshop_STRING
BestSell.Source = "SELECT ProductID, ProductName FROM Products ORDER BY Sell DESC"
BestSell.CursorType = 0
BestSell.CursorLocation = 2
BestSell.LockType = 3
BestSell.Open()
BestSell_numRows = 0
%>
<%
set newProduct = Server.CreateObject("ADODB.Recordset")
newProduct.ActiveConnection = MM_conneshop_STRING
newProduct.Source = "SELECT ProductID, ProductName, Author, Description, sImgUrl FROM Products ORDER BY AddDate DESC"
newProduct.CursorType = 0
newProduct.CursorLocation = 2
newProduct.LockType = 3
newProduct.Open()
newProduct_numRows = 0
%>
<%
Dim commendProduct__MMColParam
commendProduct__MMColParam = "Yes"
if (Request("MM_EmptyValue") <> "") then commendProduct__MMColParam = Request("MM_EmptyValue")
%>
<%
set commendProduct = Server.CreateObject("ADODB.Recordset")
commendProduct.ActiveConnection = MM_conneshop_STRING
commendProduct.Source = "SELECT ProductID, ProductName, Author, Description, sImgUrl FROM Products WHERE Commend = " + Replace(commendProduct__MMColParam, "'", "''") + " ORDER BY AddDate DESC"
commendProduct.CursorType = 0
commendProduct.CursorLocation = 2
commendProduct.LockType = 3
commendProduct.Open()
commendProduct_numRows = 0
%>
<%
Dim hotdealProduct__MMColParam
hotdealProduct__MMColParam = "Yes"
if (Request("MM_EmptyValue") <> "") then hotdealProduct__MMColParam = Request("MM_EmptyValue")
%>
<%
set hotdealProduct = Server.CreateObject("ADODB.Recordset")
hotdealProduct.ActiveConnection = MM_conneshop_STRING
hotdealProduct.Source = "SELECT ProductID, ProductName, Author, Description, Price, ListPrice, sImgUrl FROM Products WHERE HotDeal = " + Replace(hotdealProduct__MMColParam, "'", "''") + " ORDER BY AddDate DESC"
hotdealProduct.CursorType = 0
hotdealProduct.CursorLocation = 2
hotdealProduct.LockType = 3
hotdealProduct.Open()
hotdealProduct_numRows = 0
%>
<%
Dim rsCounter__MMColParam
rsCounter__MMColParam = "1"
if (Request("MM_EmptyValue") <> "") then rsCounter__MMColParam = Request("MM_EmptyValue")
%>
<%
set rsCounter = Server.CreateObject("ADODB.Recordset")
rsCounter.ActiveConnection = MM_conneshop_STRING
rsCounter.Source = "SELECT * FROM counter WHERE ID = " + Replace(rsCounter__MMColParam, "'", "''") + ""
rsCounter.CursorType = 0
rsCounter.CursorLocation = 2
rsCounter.LockType = 3
rsCounter.Open()
rsCounter_numRows = 0
dd= rsCounter.Fields.Item("digit").Value 
counter=rsCounter.Fields.Item("countnum").Value 
cc = Len(counter) 
c01 = Replace(counter,"1","<img src=""images/counter/1.gif"">")
c02 = Replace(c01,"2","<img src=""images/counter/2.gif"">")
c03 = Replace(c02,"3","<img src=""images/counter/3.gif"">")
c04 = Replace(c03,"4","<img src=""images/counter/4.gif"">")
c05 = Replace(c04,"5","<img src=""images/counter/5.gif"">")
c06 = Replace(c05,"6","<img src=""images/counter/6.gif"">")
c07 = Replace(c06,"7","<img src=""images/counter/7.gif"">")
c08 = Replace(c07,"8","<img src=""images/counter/8.gif"">")
c09 = Replace(c08,"9","<img src=""images/counter/9.gif"">")
c10 = Replace(c09,"0","<img src=""images/counter/0.gif"">")
If dd > cc Then 
For i = 1 to dd 
If i = cc Then
a= dd - i
zero = ""
for x = 1 to a
zero= zero & "<img src=""images/counter/0.gif"">"
Next
fullcounter = zero & c10 
End If
Next
counterpic= fullcounter
Else
counterpic= c10 
End If
%>
<%

set Command1 = Server.CreateObject("ADODB.Command")
Command1.ActiveConnection = MM_conneshop_STRING
Command1.CommandText = "UPDATE counter  SET countnum = countnum + 1  WHERE ID = 1"
Command1.CommandType = 1
Command1.CommandTimeout = 0
Command1.Prepared = true
Command1.Execute()

%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
booksub_numRows = booksub_numRows + Repeat1__numRows
%>
<%
Dim Repeat2__numRows
Repeat2__numRows = -1
Dim Repeat2__index
Repeat2__index = 0
videosub_numRows = videosub_numRows + Repeat2__numRows
%>
<%
Dim Repeat3__numRows
Repeat3__numRows = -1
Dim Repeat3__index
Repeat3__index = 0
musicsub_numRows = musicsub_numRows + Repeat3__numRows
%>
<%
Dim Repeat4__numRows
Repeat4__numRows = 10
Dim Repeat4__index
Repeat4__index = 0
HotProduct_numRows = HotProduct_numRows + Repeat4__numRows
%>
<%
Dim Repeat5__numRows
Repeat5__numRows = 10
Dim Repeat5__index
Repeat5__index = 0
BestSell_numRows = BestSell_numRows + Repeat5__numRows
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
<SCRIPT RUNAT=SERVER LANGUAGE=VBSCRIPT>	
function DoTrimProperly(str, nNamedFormat, properly, pointed, points)
  dim strRet
  strRet = Server.HTMLEncode(str)
  strRet = replace(strRet, vbcrlf,"")
  strRet = replace(strRet, vbtab,"")
  If (LEN(strRet) > nNamedFormat) Then
    strRet = LEFT(strRet, nNamedFormat)			
    If (properly = 1) Then					
      Dim TempArray								
      TempArray = split(strRet, " ")	
      Dim n
      strRet = ""
      for n = 0 to Ubound(TempArray) - 1
        strRet = strRet & " " & TempArray(n)
      next
    End If
    If (pointed = 1) Then
      strRet = strRet & points
    End If
  End If
  DoTrimProperly = strRet
End Function
</SCRIPT>
<html>
<head>
<title>网上商城</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="style.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" topmargin="2">
<table width="760" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td background="images/topback.gif" width="130"><img src="images/sitelogo.gif" width="130" height="88"></td>
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
<form name="search_form" method="get" action="quick_search.asp">
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
            <td width="50%">　<%=(counterpic)%></td>
            <td width="50%" valign="middle" align="center"> 
              <select name="mnuCategory">
                <option value="1" selected>在图书商城中</option>
                <option value="2">在影视商城中</option>
                <option value="3">在音乐商城中</option>
              </select>
              <input type="text" name="textPname" size="20" maxlength="50">
              <input type="submit" name="Submit" value="搜索">
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</form>
<table width="760" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td width="160" valign="top"> 
      <table width="160" border="0" cellspacing="1" cellpadding="0" bgcolor="#000000">
        <tr> 
          <td bgcolor="#006699" height="22" valign="middle"> 
            <div align="center"><font color="#FFFFFF"><img src="images/icon_arrow_d.gif" width="14" height="14">　产品分类　<img src="images/icon_arrow_d.gif" width="14" height="14"></font></div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#FFFFFF"> 
            <table width="100%" border="0" cellspacing="2" cellpadding="2">
              <tr> 
                <td valign="middle" align="center"><a href="category.asp?CategoryID=1"><img src="images/shop_book.gif" width="81" height="15" border="0"></a></td>
              </tr>
              <tr> 
                <td align="center" bgcolor="#FFFF99"> 
                  <% 
While ((Repeat1__numRows <> 0) AND (NOT booksub.EOF)) 
%>
                  <table width="100%" border="0" cellspacing="0" cellpadding="2">
                    <tr> 
                      <td valign="middle" align="center" width="25%"><img src="images/board_arrow.gif" width="17" height="13"></td>
                      <td valign="middle" align="left" width="75%"><A HREF="subcategory.asp?<%= MM_keepURL & MM_joinChar(MM_keepURL) & "SubCategoryID=" & booksub.Fields.Item("SubCategoryID").Value %>"><%=(booksub.Fields.Item("SubCategoryName").Value)%></A></td>
                    </tr>
                  </table>
                  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  booksub.MoveNext()
Wend
%>
                </td>
              </tr>
              <tr> 
                <td valign="middle" align="center"><a href="category.asp?CategoryID=2"><img src="images/shop_video.gif" width="81" height="19" border="0"></a></td>
              </tr>
              <tr> 
                <td align="center" bgcolor="#FF99FF"> 
                  <% 
While ((Repeat2__numRows <> 0) AND (NOT videosub.EOF)) 
%>
                  <table width="100%" border="0" cellspacing="0" cellpadding="2">
                    <tr> 
                      <td width="25%" align="center" valign="middle"><img src="images/board_arrow.gif" width="17" height="13"></td>
                      <td width="75%" align="left" valign="middle"><A HREF="subcategory.asp?<%= MM_keepURL & MM_joinChar(MM_keepURL) & "SubCategoryID=" & videosub.Fields.Item("SubCategoryID").Value %>" class="white"><%=(videosub.Fields.Item("SubCategoryName").Value)%></A></td>
                    </tr>
                  </table>
                  <% 
  Repeat2__index=Repeat2__index+1
  Repeat2__numRows=Repeat2__numRows-1
  videosub.MoveNext()
Wend
%>
                </td>
              </tr>
              <tr> 
                <td valign="middle" align="center"><a href="category.asp?CategoryID=3"><img src="images/shop_music.gif" width="81" height="20" border="0"></a></td>
              </tr>
              <tr> 
                <td align="center" bgcolor="#99CCFF"> 
                  <% 
While ((Repeat3__numRows <> 0) AND (NOT musicsub.EOF)) 
%>
                  <table width="100%" border="0" cellspacing="0" cellpadding="2">
                    <tr> 
                      <td width="25%" align="center" valign="middle"><img src="images/board_arrow.gif" width="17" height="13"></td>
                      <td width="75%" align="left" valign="middle"><A HREF="subcategory.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "SubCategoryID=" & musicsub.Fields.Item("SubCategoryID").Value %>" class="yellow"><%=(musicsub.Fields.Item("SubCategoryName").Value)%></A></td>
                    </tr>
                  </table>
                  <% 
  Repeat3__index=Repeat3__index+1
  Repeat3__numRows=Repeat3__numRows-1
  musicsub.MoveNext()
Wend
%>
                </td>
              </tr>
              <tr> 
                <td valign="middle" align="center">&nbsp;</td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
    <td width="420" valign="top" align="center"> <br>
      <table width="96%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="images/index_newproduct.gif" width="303" height="19"></td>
        </tr>
        <tr> 
          <td> 
            <table width="100%" border="0" cellspacing="2" cellpadding="2">
              <tr> 
                <td rowspan="3" width="25%" align="center"><A HREF="product.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "ProductID=" & newProduct.Fields.Item("ProductID").Value %>"><img src="images%5Cproduct%5C<%=(newProduct.Fields.Item("sImgUrl").Value)%>" border="0"></A></td>
                <td width="75%"><A HREF="product.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "ProductID=" & newProduct.Fields.Item("ProductID").Value %>" class="productName"><%=(newProduct.Fields.Item("ProductName").Value)%></A></td>
              </tr>
              <tr> 
                <td width="75%" class="a"><%=(newProduct.Fields.Item("Author").Value)%></td>
              </tr>
              <tr> 
                <td width="75%"> 
                  <% =(DoTrimProperly((newProduct.Fields.Item("Description").Value), 100, 0, 1, "...")) %>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td valign="bottom" align="right">&nbsp;</td>
        </tr>
      </table>
      <br>
      <table width="96%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td align="right"><img src="images/index_commend.gif" width="303" height="19"></td>
        </tr>
        <tr> 
          <td> 
            <table width="100%" border="0" cellspacing="2" cellpadding="2" height="52">
              <tr> 
                <td width="75%" align="right"><A HREF="product.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "ProductID=" & commendProduct.Fields.Item("ProductID").Value %>" class="productName"><%=(commendProduct.Fields.Item("ProductName").Value)%></A></td>
                <td rowspan="3" align="center"><A HREF="product.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "ProductID=" & commendProduct.Fields.Item("ProductID").Value %>"><img src="images%5Cproduct%5C<%=(commendProduct.Fields.Item("sImgUrl").Value)%>" border="0"></A> 
                </td>
              </tr>
              <tr> 
                <td width="75%" align="right"><%=(commendProduct.Fields.Item("Author").Value)%></td>
              </tr>
              <tr> 
                <td width="75%"> 
                  <% =(DoTrimProperly((commendProduct.Fields.Item("Description").Value), 100, 0, 1, "....")) %>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
      </table>
      <br>
      <table width="96%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="images/index_hotprice.gif" width="303" height="19"></td>
        </tr>
        <tr> 
          <td> 
            <table width="100%" border="0" cellspacing="2" cellpadding="2">
              <tr> 
                <td rowspan="3" width="18%" align="center"><A HREF="product.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "ProductID=" & hotdealProduct.Fields.Item("ProductID").Value %>"><img src="images%5Cproduct%5C<%=(hotdealProduct.Fields.Item("sImgUrl").Value)%>" border="0"></A></td>
                <td colspan="2"><A HREF="product.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "ProductID=" & hotdealProduct.Fields.Item("ProductID").Value %>" class="productName"><%=(hotdealProduct.Fields.Item("ProductName").Value)%></A></td>
              </tr>
              <tr> 
                <td colspan="2"><%=(hotdealProduct.Fields.Item("Author").Value)%></td>
              </tr>
              <tr> 
                <td width="20%">原价：<span class="hotPrice"><%=(hotdealProduct.Fields.Item("Price").Value)%></span>元</td>
                <td width="80%">现价：<%=(hotdealProduct.Fields.Item("ListPrice").Value)%>元</td>
              </tr>
              <tr> 
                <td colspan="3"> 
                  <% =(DoTrimProperly((hotdealProduct.Fields.Item("Description").Value), 100, 0, 1, "...")) %>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td valign="bottom" align="right">&nbsp;</td>
        </tr>
      </table>
    </td>
    <td width="180" valign="top"> 
      <table width="180" border="0" cellspacing="1" cellpadding="0" bgcolor="#000000">
        <tr> 
          <td bgcolor="#006699" height="22" valign="middle"> 
            <div align="center"><font color="#FFFFFF"><img src="images/nav_document.gif" width="16" height="16"> 
              今日热点</font></div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#FFFFFF" align="center" valign="top"> 
            <table width="100%" border="0" cellspacing="2" cellpadding="2">
              <tr> 
                <td> 
                  <% 
While ((Repeat4__numRows <> 0) AND (NOT HotProduct.EOF)) 
%>
                  <table width="100%" border="0" cellspacing="0" cellpadding="2">
                    <tr> 
                      <td width="10%" align="center" valign="top"><img src="images/board_arrow_u.gif" width="17" height="13"></td>
                      <td width="90%" align="left" valign="middle"><A HREF="product.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "ProductID=" & HotProduct.Fields.Item("ProductID").Value %>"><%=(HotProduct.Fields.Item("ProductName").Value)%></A></td>
                    </tr>
                  </table>
                  <% 
  Repeat4__index=Repeat4__index+1
  Repeat4__numRows=Repeat4__numRows-1
  HotProduct.MoveNext()
Wend
%>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <br>
      <table width="180" border="0" cellspacing="1" cellpadding="0" bgcolor="#000000">
        <tr> 
          <td bgcolor="#FFCC66" height="22" valign="middle"> 
            <div align="center"><font color="#FFFFFF"><img src="images/icon_arrow_u.gif" width="14" height="14">　</font>销售排行　<img src="images/icon_arrow_u.gif" width="14" height="14"></div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#FFFFFF"> 
            <table width="100%" border="0" cellspacing="2" cellpadding="2">
              <tr> 
                <td> 
                  <% 
While ((Repeat5__numRows <> 0) AND (NOT BestSell.EOF)) 
%>
                  <table width="100%" border="0" cellspacing="0" cellpadding="2">
                    <tr> 
                      <td width="10%" align="center" valign="top"><img src="images/board_arrow_u.gif" width="17" height="13"></td>
                      <td width="90%" align="left" valign="middle"><A HREF="product.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "ProductID=" & BestSell.Fields.Item("ProductID").Value %>" class="red"><%=(BestSell.Fields.Item("ProductName").Value)%></A></td>
                    </tr>
                  </table>
                  <% 
  Repeat5__index=Repeat5__index+1
  Repeat5__numRows=Repeat5__numRows-1
  BestSell.MoveNext()
Wend
%>
                </td>
              </tr>
            </table>
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
booksub.Close()
%>
<%
videosub.Close()
%>
<%
musicsub.Close()
%>
<%
HotProduct.Close()
%>
<%
BestSell.Close()
%>
<%
newProduct.Close()
%>
<%
commendProduct.Close()
%>
<%
hotdealProduct.Close()
%>
<%
rsCounter.Close()
%>
