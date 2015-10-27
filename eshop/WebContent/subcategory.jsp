<%@ page language="java" contentType="text/html; charset=gb2312"
	pageEncoding="gb2312"  %>

<%@ taglib uri="http://java.sun.com/jsp/jstl/core" prefix="c" %>

<html>
<head>
<title>网上商城</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="style/${subCategory.style}" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" topmargin="2">
<table width="760" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td background="images/topback.gif" width="130"><img src="images/sitelogo.gif" height="88"></td>
    <td background="images/topback.gif" width="500" align="center" valign="middle"><a href="http://www.fans8.com" target="_blank"><img src="images/fans8.gif" width="468" height="60" border="0"></a> 
    </td>
    <td background="images/topback.gif" width="130"> 
	<table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td valign="middle" align="center"><a href="cart.asp"><img src="images/button_cart.gif" width="87" height="18" border="0"></a></td>
        </tr>
        <tr> 
          <td valign="middle" align="center"><a href="checkorder_login.asp"><img src="images/button_ddcx.gif" width="87" height="18" border="0"></a></td>
        </tr>
        <tr> 
          
    <td valign="middle" align="center"><a href="customer_register.asp"><img src="images/button_regist.gif" width="87" height="18" border="0"></a></td>
        </tr>
      </table>
      </td>
  </tr>
</table>
<form name="search_form" method="get" action="quick_search.asp">
  <table width="760" border="0" cellspacing="1" cellpadding="0" align="center" bgcolor="#000000">
    <tr> 
      <td bgcolor="#FF9900" height="22" valign="middle" align="center"> 
	  <table width="80%" border="0" cellspacing="2" cellpadding="2">
          <tr align="center" valign="middle"> 
            <td><a href="newproduct.asp" class="white">新品快递</a></td>
            
    <td><a href="commend.asp" class="white">重点推荐</a></td>
            
    <td><a href="bestsell.asp" class="white">销售排行</a></td>
            
    <td><a href="bestprice.asp" class="white">特价商品</a></td>
          </tr>
        </table>
        </td>
    </tr>
    <tr> 
      <td bgcolor="#FFCC66" height="22"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="50%">　<a href="index.jsp" class="red">首页</a> &gt; <a href="category.jsp?id=" class="red">${subCategory.category.name}</a> &gt; ${subCategory.name}</td>
            <td width="50%" align="center" valign="middle"> 
              <select name="mnuCategory">
                <option value="${CategoryID}" selected>在本类商城中</option>
                <option value="1">在图书商城中`</option>
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
    <td width="75%" valign="top" align="left"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="49%"><img src="images/subcategory/${subCategory.img}"></td>
          <td width="51%" valign="bottom" align="center"> 
            <form name="form1" method="get" action="subcategory.jsp">
              <input type="hidden" name="SubCategoryID">
              商品按 
              <select name="order" class="list">
                <option value="AddDate" selected >上架时间</option>
                <option value="ListPrice">当前价格</option>
                <option value="PubDate">出版时间</option>
                <option value="Sell">销售数量</option>
                <option value="Visits">点击次数</option>
              </select>
              以 
              <select name="sort" class="list">
                <option value="DESC"  selected>降序</option>
                <option value="ASC" >升序</option>
              </select>
              <input type="submit" value="排序">
            </form>
          </td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td valign="middle" align="left">当前为第 ${subProducts_first} 到 ${subProducts_last} 条记录，共${subProducts_total}条商品信息 </td>
        </tr>
      </table>
      
      
      <table width="99%" border="0" cellspacing="2" cellpadding="5">
        <tr> 
          <td rowspan="3" width="20%" valign="top" align="center">
          <a href="product.asp?"><img src="images/product/${sImgUrl}" border="0"/></a>
          </td>
          <td width="80%">
          <a href="product.asp?" class="productName">${ProductName}</a>
           
            <!-- 判断是否显示折扣图标开始 -->
            <img src="images/hotprice.gif" width="24" height="24"> 
            <!-- 判断是否显示折扣图标结束-->
            
            　<a href="#"><img border=0 src="images/addtocart.gif" width="30" height="18" alt="添加到购物车"></A></td>
        </tr>
        <tr> 
          <td width="80%"> 
            <!-- 是否显示原价：条件是price not eq 0 -->
            原价：<span class="hotPrice">${Price}</span>元　现价： 
            <!-- 是否显示原价 -->
            
           <!-- 是否显示价格：条件是price eq 0 -->
            价格： 
            <!-- 是否显示价格：条件是price eq 0 -->
          ${ListPrice}元　</td>
        </tr>
        <tr> 
          <td width="80%"> 
            ${Description}
          </td>
        </tr>
        <tr> 
          <td colspan="2" height="1" bgcolor="#CCCCCC"></td>
        </tr>
      </table>
      
      <table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td> 当前为第 ${subProducts_first} 到 ${subProducts_last} 条记录，共${subProducts_total}条商品信息 </td>
        </tr>
        <tr> 
          <td align="center"> 
            <table border="0" width="90%">
              <tr> 
                <td width="15%" align="left"> 
                  <!--是否显示第一页 -->
                  <a href="#" class="navi">第一页</a> 
                  <!--是否显示第一页 -->
                </td>
                <td width="15%" align="left"> 
                  <!--是否显示前一页 -->
                  <a href="#" class="navi">前一页</a> 
                  <!--是否显示前一页 -->
                </td>
                <td width="40%" align="center"> 
                  
                </td>
                <td width="15%" align="right"> 
                  <!--是否显示下一页 -->
                  <a href="#" class="navi">下一页</a> 
                  <!--是否显示下一页 -->
                </td>
                <td width="15%" align="right"> 
                  <!--是否显示最末页 -->
                  <a href="#" class="navi">最末页</a> 
                  <!--是否显示最末页 -->
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
    <td width="25%" valign="top" align="right"> 
      <table width="100%" border="0" cellspacing="2" cellpadding="2" class="commendbox" >
        <tr valign="top"> 
          <td align="center" valign="middle" colspan="2" height="24" class="commendtitle">.:: 
            本类推荐商品 ::.</td>
        </tr>
        
        <tr> 
          <td width="8%" height="17" valign="top"><img src="images/category/square.gif" width="9" height="9"></td>
          <td height="17"><a href="product.asp?" class="commend">${ProductName}</a>
		  </td>
        </tr>
        
        <tr> 
          <td colspan="2">&nbsp;</td>
        </tr>
      </table>
      <br>
      <table width="100%" border="0" cellspacing="2" cellpadding="2" class="bestsellbox" >
        <tr valign="top"> 
          <td align="center" valign="middle" colspan="2" height="24" class="bestselltitle">.:: 
            本类畅销排行 ::.</td>
        </tr>
        
        <tr> 
          <td width="8%" height="17" valign="top"><img src="images/category/square.gif" width="9" height="9"></td>
          <td height="17"><a href="product.asp?" class="bestsell">${ProductName}</a></td>
        </tr>
        
        <tr> 
          <td colspan="2">&nbsp;</td>
        </tr>
      </table>
      <br>
      <table width="100%" border="0" cellspacing="2" cellpadding="2" class="commendbox" >
        <tr> 
          <td align="center" valign="middle" height="24" class="commendtitle"> 
            .:: 本类组合搜索 ::.</td>
        </tr>
        <tr> 
          <td> 
            <form name="comp_form" method="get" action="comp_search.asp">
              <table width="100%" border="0" cellspacing="2" cellpadding="2">
                <tr> 
                  <td width="54%" align="right" valign="middle">商品名:</td>
                  <td width="46%" align="left" valign="middle"> 
                    <input type="text" name="tex_name" size="15">
                  </td>
                </tr>
                <tr> 
                  <td width="54%" align="right" valign="middle">作者/演员:</td>
                  <td width="46%" align="left" valign="middle"> 
                    <input type="text" name="tex_author" size="15">
                  </td>
                </tr>
                <tr> 
                  <td width="54%" align="right" valign="middle">出版单位:</td>
                  <td width="46%" align="left" valign="middle"> 
                    <input type="text" name="tex_supply" size="15">
                  </td>
                </tr>
                <tr> 
                  <td width="54%" align="right" valign="middle">出版日期:</td>
                  <td width="46%" align="left" valign="middle"> 
                    <select name="menu_pub">
                      <option value="1" selected>所有出版日期</option>
                      <option value="date()-PubDate&lt;7">本周出版发行</option>
                      <option value="date()-PubDate&lt;15">半个月内出版</option>
                      <option value="date()-PubDate&gt;30">一个月前出版</option>
                      <option value="date()-PubDate&gt;90">三个月前出版</option>
                      <option value="date()-PubDate&gt;180">半年前出版</option>
                      <option value="date()-PubDate&gt;365">一年前出版</option>
                    </select>
                  </td>
                </tr>
                <tr> 
                  <td width="54%" align="right" valign="middle">折扣范围:</td>
                  <td width="46%" align="left" valign="middle"> 
                    <select name="menu_hotdeal">
                      <option value="1" selected>无论打折与否</option>
                      <option value="HotDeal=YES">所有打折的</option>
                      <option value="HotDeal=NO">所有不打折的</option>
                    </select>
                  </td>
                </tr>
                <tr> 
                  <td width="54%" align="right" valign="middle">售价范围:</td>
                  <td width="46%" align="left" valign="middle"> 
                    <select name="menu_price">
                      <option value="1" selected>所有价格</option>
                      <option value="ListPrice&lt;5">低于5元的</option>
                      <option value="ListPrice&lt;10">低于10元的</option>
                      <option value="ListPrice&lt;30">低于30元的</option>
                      <option value="ListPrice&gt;50">高于50元的</option>
                      <option value="ListPrice&gt;100">高于100元的</option>
                    </select>
                  </td>
                </tr>
                <tr valign="middle"> 
                  <td colspan="2" align="center"> 
                    <input type="hidden" name="h_subcateid" value="${SubCategoryID}">
                    <input type="submit" value="开始搜索">
                    <input type="reset" value="清除重填" name="submit2">
                  </td>
                </tr>
              </table>
            </form>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<br>

<table width="760" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td background="images/topback.gif" align="center" height="16"><font color="#FFFFFF">copyright 
      2001 Powered by Peter.HJ</font></td>
  </tr>
</table>
</body>
</html>
