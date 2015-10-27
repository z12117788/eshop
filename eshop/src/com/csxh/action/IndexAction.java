package com.csxh.action;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import javax.servlet.ServletContext;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpSession;

import com.csxh.model.Category;
import com.csxh.model.Product;
import com.csxh.model.SubCategory;
import com.csxh.util.JdbcUtil;

public class IndexAction {

	protected HttpServletRequest request;

	public void setRequest(HttpServletRequest request) {
		this.request = request;

	}

	public void setSession(HttpSession session) {
		// TODO Auto-generated method stub

	}

	public void setApplication(ServletContext servletContext) {
		// TODO Auto-generated method stub

	}

	public void setParams(Map<String, String[]> parameterMap) {
		// TODO Auto-generated method stub

	}

	public String handle() {

		// 获取index.jsp所需要的模型数据，并保存到jsp页的内置对象中

		// 获取大类别及其子类别列表
		String result = "fail";

		List<Category> categoryList = JdbcUtil.queryObjectList(Category.class);

		if (categoryList != null && categoryList.size() > 0) {

			for (Category c : categoryList) {
				List<SubCategory> subList = JdbcUtil.queryObjectList(SubCategory.class, "categoryId=" + c.getId());
				c.setChildren(subList);
			}

			this.request.setAttribute("categoryList", categoryList);

			result = "success";
		}

		// 获取今日热点产品列表
		List<Product> hotProductList = new ArrayList<Product>(10);

		List<Object[]> objectList = JdbcUtil
				.queryForObjectList("select top 10 id,name from product order by visits desc");
		for (Object[] o : objectList) {

			Product p = new Product();

			p.setId((String) o[0]);
			p.setName((String) o[1]);

			hotProductList.add(p);

		}

		if (hotProductList.size() > 0) {
			this.request.setAttribute("hotProductList", hotProductList);
			result = "success";
		}

		// 获取销售排行列表
		List<Product> sellProductList = new ArrayList<Product>(10);

		objectList = JdbcUtil.queryForObjectList("select top 10 id,name from product order by sell desc");
		for (Object[] o : objectList) {

			Product p = new Product();

			p.setId((String) o[0]);
			p.setName((String) o[1]);

			sellProductList.add(p);

		}

		if (sellProductList.size() > 0) {
			this.request.setAttribute("sellProductList", sellProductList);
			result = "success";
		}

		// 获取最新商品
		List<Product> newProductList = new ArrayList<Product>(10);
		
		objectList = JdbcUtil.queryForObjectList(
				"select top 1 id,name,author,smallImg,description from product order by addDate desc");
		for (Object[] o : objectList) {

			Product p = new Product();

			p.setId((String) o[0]);
			p.setName((String) o[1]);
			p.setAuthor((String) o[2]);
			p.setSmallImg((String) o[3]);
			p.setDescription((String) o[4]);

			newProductList.add(p);
			break;

		}

		if (newProductList.size() > 0) {
			this.request.setAttribute("newProductList", newProductList);
			result = "success";
		}

		// 获取推荐商品
		List<Product> commendProductList = new ArrayList<Product>(10);
		
		objectList = JdbcUtil.queryForObjectList(
				"select top 1 id,name,author,smallImg,description from product where commend=1");
		for (Object[] o : objectList) {
			
			Product p = new Product();
			
			p.setId((String) o[0]);
			p.setName((String) o[1]);
			p.setAuthor((String) o[2]);
			p.setSmallImg((String) o[3]);
			p.setDescription((String) o[4]);
			
			commendProductList.add(p);
			break;
			
		}
		
		if (commendProductList.size() > 0) {
			this.request.setAttribute("commendProductList", commendProductList);
			result = "success";
		}

		// 获取折扣商品
		List<Product> discountProductList = new ArrayList<Product>(10);
		
		objectList = JdbcUtil.queryForObjectList(
				"select top 1 id,name,author,smallImg,description,price,listPrice from product order by listPrice desc ");
		for (Object[] o : objectList) {
			
			Product p = new Product();
			
			p.setId((String) o[0]);
			p.setName((String) o[1]);
			p.setAuthor((String) o[2]);
			p.setSmallImg((String) o[3]);
			p.setDescription((String) o[4]);
			p.setPrice((BigDecimal) o[5]);
			p.setListPrice((BigDecimal) o[6]);
			
			discountProductList.add(p);
			break;
			
		}
		
		if (discountProductList.size() > 0) {
			this.request.setAttribute("discountProductList", discountProductList);
			result = "success";
		}

		return result;
	}

}
