<%@ page language="java" contentType="text/html; charset=gb2312"
	pageEncoding="gb2312"  %>

<%@ taglib uri="http://java.sun.com/jsp/jstl/core" prefix="c" %>
	
<html>
<head>
<title>网上商城</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="style.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" topmargin="2">
<table width="760" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td width="130" background="images/topback.gif"><img src="images/sitelogo.gif" width="130" height="88"></td>
    <td background="images/topback.gif" width="500" align="center" valign="middle">

<table width="468" border="0" cellspacing="0" cellpadding="0">
  
  <tr> 
    
    <td align="center" valign="middle"><a href="#" target="_blank"><img src="adbanner/fans8.gif" width="468" height="60" border="0" alt=".:: 狂迷俱乐部 ::."></a></td>
    
  </tr>
</table>
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
            <td width="50%">　${counterpic}</td>
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
          <!--开始输出大类别数据 -->
            <table width="100%" border="0" cellspacing="2" cellpadding="2">
              
              <!-- 使用EL表达式(${""})和JSTL标签来代替代码标签  -->
              
              <c:forEach items="${categoryList}" var="category" >
              
	              <tr> 
	                <td valign="middle" align="center"><a href="category.jsp">
	                
	                <c:if test="${category.id==1}">	                
	                    <img src="images/${category.smallImg}" width="81" height="15" border="0">	                   
	                </c:if>
	                
	                <c:if test="${category.id>1}">	                
	                    <img src="images/${category.smallImg}" width="81" height="20" border="0">
	                </c:if>
	                
	                </a>
	                </td>
	              </tr>
	              <tr> 
	                <td align="center" bgcolor="#FFFFE8"> 
	                  
	                  <table width="100%" border="0" cellspacing="0" cellpadding="2">
	                  
                      <!-- 商品子类开始 -->
	                  <c:forEach items="${category.children }" var="subcategory">
	                  
	                    <tr> 
	                      <td valign="middle" align="center" width="25%">
	                         <img src="images/board_arrow.gif" width="17" height="13">
	                      </td>
	                      <td valign="middle" align="left" width="75%">
	                      <a href="subcategory.jsp?id=${subcategory.id}" class="white">${subcategory.name}</a>
	                      </td>
	                    </tr>
	                  
	                  </c:forEach>
                       <!-- 商品子类结束 -->
	                  
	                  </table>
	                  </td>
	              </tr>
              
              
              </c:forEach>
              
             
              <tr> 
                <td valign="middle" align="center">&nbsp;</td>
              </tr>
            </table>
            <!--结束输出大类别数据 -->
            
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
          
          <!-- 最新产品开始 -->
          
            <table width="100%" border="0" cellspacing="2" cellpadding="2">
              <c:forEach items="${newProductList}" var="newProduct">
              
              <tr> 
                <td rowspan="3" width="25%" align="center">
                	<a href="product.jsp?"><img src="images/product/${newProduct.smallImg}" border="0"></a>
                </td>
                <td width="75%">
                	<a href="product.jsp?" class="productName">${newProduct.name}</a>
                </td>
              </tr>
              <tr> 
                <td width="75%" class="a">${newProduct.author}</td>
              </tr>
              <tr> 
                <td width="75%">${newProduct.description}</td>
              </tr>

              </c:forEach>
              
            </table>
            <!-- 最新产品结束 -->
            
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
          <!--推荐商品开始 -->
            <table width="100%" border="0" cellspacing="2" cellpadding="2" height="52">
              
              <c:forEach items="${commendProductList }" var="commendProduct">
              
              <tr> 
                <td width="75%" align="right">
                	<a href="product.jsp?" class="productName">${commendProduct.name}</a>
                </td>
                <td rowspan="3" align="center">
	                <a href="product.asp?">
	                	<img src="images/product/${commendProduct.smallImg}" border="0">
	                </a> 
                </td>
              </tr>
              <tr> 
                <td width="75%" align="right">${Author}</td>
              </tr>
              <tr> 
                <td width="75%"> 
                
                  ${commendProduct.description}
                
                </td>
              </tr>
              
              </c:forEach>
              
            </table>
          </td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
      </table>
      <!--推荐商品结束 -->
      
      <br>
      <table width="96%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="images/index_hotprice.gif" width="303" height="19"></td>
        </tr>
        <tr> 
          <td> 
          <!-- 折扣商品开始-->
            <table width="100%" border="0" cellspacing="2" cellpadding="2">
              
              <c:forEach items="${discountProductList }" var="discountProduct">
              
              <tr> 
                <td rowspan="3" width="18%" align="center">
                	<a href="product.asp?"><img src="images/product/${discountProduct.smallImg}" border="0"></a>
                </td>
                <td colspan="2">
                	<a href="product.asp?" class="productName">${discountProduct.name}</a>
                </td>
              </tr>
              <tr> 
                <td colspan="2">${Author}</td>
              </tr>
              <tr> 
                <td width="20%">原价：<span class="hotPrice">${discountProduct.price}</span>元</td>
                <td width="80%">现价：${discountProduct.listPrice}元</td>
              </tr>
              <tr> 
                <td colspan="3"> 
                  ${discountProduct.description}
                </td>
              </tr>
              
              </c:forEach>
              
            </table>
            <!-- 折扣商品结束-->
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
                <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
                <!-- 今日热点开始 -->
                
                <c:forEach items="${hotProductList}" var="p">
                
                    <tr> 
                      <td width="10%" align="center" valign="top"><img src="images/board_arrow_u.gif" width="17" height="13"></td>
                      <td width="90%" align="left" valign="middle">
                      <a href="product.jsp?" class="red">${p.name}</a></td>
                    </tr>
                
                
                </c:forEach>
                
                <!-- 今日热点结束 -->
                    
                  </table></td>
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
                  
                  <table width="100%" border="0" cellspacing="0" cellpadding="2">
                  <!--销售排行开始-->
                   
                   <c:forEach items="${sellProductList}" var="sellProduct">
                   
                    <tr> 
                      <td width="10%" align="center" valign="top">
                      	<img src="images/board_arrow_u.gif" width="17" height="13">
                      </td>
                      <td width="90%" align="left" valign="middle">
                      	<a href="product.asp?" class="red">${sellProduct.name}</a>
                      </td>
                    </tr>
                    
                    </c:forEach>
                    
                    <!--销售排行开始-->
                  </table>
                  
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
<table width="760" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td background="images/topback.gif" align="center" height="16"><font color="#FFFFFF">copyright 
      2001 Powered by Peter.HJ</font></td>
  </tr>
</table>
</body>
</html>
