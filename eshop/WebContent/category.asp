<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/conneshop.asp" -->
<%
Dim subcategory__MMColParam
subcategory__MMColParam = "1"
if (Request.QueryString("CategoryID") <> "") then subcategory__MMColParam = Request.QueryString("CategoryID")
%>
<%
set subcategory = Server.CreateObject("ADODB.Recordset")
subcategory.ActiveConnection = MM_conneshop_STRING
subcategory.Source = "SELECT SubCategoryID, SubCategoryName FROM SubCategories WHERE CategoryID = " + Replace(subcategory__MMColParam, "'", "''") + ""
subcategory.CursorType = 0
subcategory.CursorLocation = 2
subcategory.LockType = 3
subcategory.Open()
subcategory_numRows = 0
%>
<%
Dim cate__MMColParam
cate__MMColParam = "1"
if (Request.QueryString("CategoryID") <> "") then cate__MMColParam = Request.QueryString("CategoryID")
%>
<%
set cate = Server.CreateObject("ADODB.Recordset")
cate.ActiveConnection = MM_conneshop_STRING
cate.Source = "SELECT * FROM Categories WHERE CategoryID = " + Replace(cate__MMColParam, "'", "''") + ""
cate.CursorType = 0
cate.CursorLocation = 2
cate.LockType = 3
cate.Open()
cate_numRows = 0
%>
<%
Dim BestSell__MMColParam
BestSell__MMColParam = "1"
if (Request.QueryString("CategoryID") <> "") then BestSell__MMColParam = Request.QueryString("CategoryID")
%>
<%
set BestSell = Server.CreateObject("ADODB.Recordset")
BestSell.ActiveConnection = MM_conneshop_STRING
BestSell.Source = "SELECT ProductID, ProductName, Author, Description, sImgUrl FROM Products WHERE CategoryID = " + Replace(BestSell__MMColParam, "'", "''") + " ORDER BY Sell DESC"
BestSell.CursorType = 0
BestSell.CursorLocation = 2
BestSell.LockType = 3
BestSell.Open()
BestSell_numRows = 0
%>
<%
Dim newProduct__MMColParam
newProduct__MMColParam = "1"
if (Request.QueryString("CategoryID") <> "") then newProduct__MMColParam = Request.QueryString("CategoryID")
%>
<%
set newProduct = Server.CreateObject("ADODB.Recordset")
newProduct.ActiveConnection = MM_conneshop_STRING
newProduct.Source = "SELECT ProductID, ProductName, Author, Description, sImgUrl FROM Products WHERE CategoryID = " + Replace(newProduct__MMColParam, "'", "''") + " ORDER BY AddDate DESC"
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
Dim commendProduct__MMColParam2
commendProduct__MMColParam2 = "1"
if (Request.QueryString("CategoryID") <> "") then commendProduct__MMColParam2 = Request.QueryString("CategoryID")
%>
<%
set commendProduct = Server.CreateObject("ADODB.Recordset")
commendProduct.ActiveConnection = MM_conneshop_STRING
commendProduct.Source = "SELECT ProductID, ProductName, Author, Description, sImgUrl  FROM Products   WHERE Commend = " + Replace(commendProduct__MMColParam, "'", "''") + " AND CategoryID=" + Replace(commendProduct__MMColParam2, "'", "''") + "  ORDER BY AddDate DESC"
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
Dim hotdealProduct__MMColParam2
hotdealProduct__MMColParam2 = "1"
if (Request.QueryString("CategoryID") <> "") then hotdealProduct__MMColParam2 = Request.QueryString("CategoryID")
%>
<%
set hotdealProduct = Server.CreateObject("ADODB.Recordset")
hotdealProduct.ActiveConnection = MM_conneshop_STRING
hotdealProduct.Source = "SELECT ProductID, ProductName, Author, Description, Price, ListPrice, sImgUrl  FROM Products    WHERE HotDeal = " + Replace(hotdealProduct__MMColParam, "'", "''") + " AND CategoryID=" + Replace(hotdealProduct__MMColParam2, "'", "''") + "  ORDER BY AddDate DESC"
hotdealProduct.CursorType = 0
hotdealProduct.CursorLocation = 2
hotdealProduct.LockType = 3
hotdealProduct.Open()
hotdealProduct_numRows = 0
%>
<%
Dim HLooper1__numRows
HLooper1__numRows = -4
Dim HLooper1__index
HLooper1__index = 0
subcategory_numRows = subcategory_numRows + HLooper1__numRows
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = 3
Dim Repeat1__index
Repeat1__index = 0
newProduct_numRows = newProduct_numRows + Repeat1__numRows
%>
<%
Dim Repeat2__numRows
Repeat2__numRows = 3
Dim Repeat2__index
Repeat2__index = 0
hotdealProduct_numRows = hotdealProduct_numRows + Repeat2__numRows
%>
<%
Dim Repeat3__numRows
Repeat3__numRows = 3
Dim Repeat3__index
Repeat3__index = 0
commendProduct_numRows = commendProduct_numRows + Repeat3__numRows
%>
<%
Dim Repeat4__numRows
Repeat4__numRows = 3
Dim Repeat4__index
Repeat4__index = 0
BestSell_numRows = BestSell_numRows + Repeat4__numRows
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
<link rel="stylesheet" href="style/<%=(cate.Fields.Item("CategoryStyle").Value)%>" type="text/css">
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
            <td width="50%" height="30">　<a href="default.asp" class="red">首页</a> 
              &gt; <%=(cate.Fields.Item("CategoryName").Value)%></td>
            <td width="50%" valign="middle" height="30" align="center"> 
              <select name="mnuCategory">
                <option value="<%=(cate.Fields.Item("CategoryID").Value)%>" selected>在本类商城中</option>
                <option value="1">在图书商城中</option>
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
    <td width="360" valign="middle"> 
      <table>
        <%
startrw = 0
endrw = HLooper1__index
numberColumns = 4
numrows = -1
while((numrows <> 0) AND (Not subcategory.EOF))
	startrw = endrw + 1
	endrw = endrw + numberColumns
 %>
        <tr  valign="top"> 
          <%
While ((startrw <= endrw) AND (Not subcategory.EOF))
%>
          <td> <img src="images/category/square.gif" width="9" height="9"> <A HREF="subcategory.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "SubCategoryID=" & subcategory.Fields.Item("SubCategoryID").Value %>" class="subcate"><%=(subcategory.Fields.Item("SubCategoryName").Value)%></A> </td>
          <%
	startrw = startrw + 1
	subcategory.MoveNext()
	Wend
	%>
        </tr>
        <%
 numrows=numrows-1
 Wend
 %>
      </table>
    </td>
    <td valign="bottom" align="right"><img src="images/category/<%=(cate.Fields.Item("CategoryImg").Value)%>"></td>
  </tr>
</table>
<table width="760" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td colspan="2">&nbsp;</td>
  </tr>
  <tr> 
    <td valign="top" align="left" width="380"> 
      <table width="98%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td class="bar1">　　新　品</td>
        </tr>
        <tr> 
          <td valign="top"> 
            <table width="100%" border="0" cellspacing="2" cellpadding="2" class="bar1content">
              <% 
While ((Repeat1__numRows <> 0) AND (NOT newProduct.EOF)) 
%>
              <tr valign="top"> 
                <td rowspan="3"><A HREF="product.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "ProductID=" & newProduct.Fields.Item("ProductID").Value %>"><img src="images/product/<%=(newProduct.Fields.Item("sImgUrl").Value)%>" border="0"></A></td>
                <td width="72%"><A HREF="product.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "ProductID=" & newProduct.Fields.Item("ProductID").Value %>" class="productName"><%=(newProduct.Fields.Item("ProductName").Value)%></A></td>
              </tr>
              <tr> 
                <td width="72%"><%=(newProduct.Fields.Item("Author").Value)%></td>
              </tr>
              <tr> 
                <td width="72%"> 
                  <% =(DoTrimProperly((newProduct.Fields.Item("Description").Value), 80, 0, 1, "...")) %>
                </td>
              </tr>
              <tr> 
                <td width="25%" height="8"></td>
                <td width="72%" height="8"></td>
              </tr>
              <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  newProduct.MoveNext()
Wend
%>
              <tr> 
                <td width="25%">&nbsp;</td>
                <td width="72%">&nbsp;</td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <br>
      <table width="98%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td class="bar2">　　打　折</td>
        </tr>
        <tr> 
          <td> 
            <table width="100%" border="0" cellspacing="2" cellpadding="2" class="bar2content">
              <% 
While ((Repeat2__numRows <> 0) AND (NOT hotdealProduct.EOF)) 
%>
              <tr valign="top"> 
                <td rowspan="3"><A HREF="product.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "ProductID=" & hotdealProduct.Fields.Item("ProductID").Value %>"><img src="images/product/<%=(hotdealProduct.Fields.Item("sImgUrl").Value)%>" border="0"></A></td>
                <td width="72%"><A HREF="product.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "ProductID=" & hotdealProduct.Fields.Item("ProductID").Value %>" class="productName"><%=(hotdealProduct.Fields.Item("ProductName").Value)%></A></td>
              </tr>
              <tr> 
                <td width="72%"><%=(hotdealProduct.Fields.Item("Author").Value)%></td>
              </tr>
              <tr> 
                <td width="72%"> 
                  <% =(DoTrimProperly((hotdealProduct.Fields.Item("Description").Value), 80, 0, 1, "...")) %>
                </td>
              </tr>
              <tr> 
                <td width="25%">&nbsp;</td>
                <td width="72%">原价：<span class="hotPrice"><%=(hotdealProduct.Fields.Item("Price").Value)%></span>元　现价：<%=(hotdealProduct.Fields.Item("ListPrice").Value)%>元</td>
              </tr>
              <tr> 
                <td width="25%" height="8"></td>
                <td width="72%" height="8"></td>
              </tr>
              <% 
  Repeat2__index=Repeat2__index+1
  Repeat2__numRows=Repeat2__numRows-1
  hotdealProduct.MoveNext()
Wend
%>
              <tr> 
                <td width="25%">&nbsp;</td>
                <td width="72%">&nbsp;</td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
    <td width="380" align="right" valign="top"> 
      <table width="98%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td class="bar2">　　推　荐</td>
        </tr>
        <tr> 
          <td valign="top"> 
            <table width="100%" border="0" cellspacing="2" cellpadding="2" class="bar2content">
              <% 
While ((Repeat3__numRows <> 0) AND (NOT commendProduct.EOF)) 
%>
              <tr valign="top"> 
                <td rowspan="3"><A HREF="product.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "ProductID=" & commendProduct.Fields.Item("ProductID").Value %>"><img src="images/product/<%=(commendProduct.Fields.Item("sImgUrl").Value)%>" border="0"></A></td>
                <td width="72%"><A HREF="product.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "ProductID=" & commendProduct.Fields.Item("ProductID").Value %>" class="productName"><%=(commendProduct.Fields.Item("ProductName").Value)%></A></td>
              </tr>
              <tr> 
                <td width="72%"><%=(commendProduct.Fields.Item("Author").Value)%></td>
              </tr>
              <tr> 
                <td width="72%"> 
                  <% =(DoTrimProperly((commendProduct.Fields.Item("Description").Value), 80, 0, 1, "...")) %>
                </td>
              </tr>
              <tr> 
                <td width="25%" height="8"></td>
                <td width="72%" height="8"></td>
              </tr>
              <% 
  Repeat3__index=Repeat3__index+1
  Repeat3__numRows=Repeat3__numRows-1
  commendProduct.MoveNext()
Wend
%>
              <tr> 
                <td width="25%">&nbsp;</td>
                <td width="72%">&nbsp;</td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <br>
      <table width="98%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td class="bar1">　　热　卖</td>
        </tr>
        <tr> 
          <td valign="top"> 
            <table width="100%" border="0" cellspacing="2" cellpadding="2" class="bar1content">
              <% 
While ((Repeat4__numRows <> 0) AND (NOT BestSell.EOF)) 
%>
              <tr valign="top"> 
                <td rowspan="3"><A HREF="product.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "ProductID=" & BestSell.Fields.Item("ProductID").Value %>"><img src="images/product/<%=(BestSell.Fields.Item("sImgUrl").Value)%>" border="0"></A></td>
                <td width="72%"><A HREF="product.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "ProductID=" & BestSell.Fields.Item("ProductID").Value %>" class="productName"><%=(BestSell.Fields.Item("ProductName").Value)%></A></td>
              </tr>
              <tr> 
                <td width="72%"><%=(BestSell.Fields.Item("Author").Value)%></td>
              </tr>
              <tr> 
                <td width="72%"> 
                  <% =(DoTrimProperly((BestSell.Fields.Item("Description").Value), 80, 0, 1, "...")) %>
                </td>
              </tr>
              <tr> 
                <td width="25%" height="8"></td>
                <td width="72%" height="8"></td>
              </tr>
              <% 
  Repeat4__index=Repeat4__index+1
  Repeat4__numRows=Repeat4__numRows-1
  BestSell.MoveNext()
Wend
%>
              <tr> 
                <td width="25%">&nbsp;</td>
                <td width="72%">&nbsp;</td>
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
subcategory.Close()
%>
<%
cate.Close()
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
