package com.csxh.model;

import java.io.Serializable;
import java.util.List;

public class Category implements Serializable {

	private Integer id;
	private String name;
	private String description;
	private String bigImg;
	private String smallImg;
	private String style;
	
	List<SubCategory> children;
	
	
	public List<SubCategory> getChildren() {
		return children;
	}
	
	public void setChildren(List<SubCategory> children) {
		this.children = children;
	}
	
	public String getBigImg() {
		return bigImg;
	}
	public void setBigImg(String bigImg) {
		this.bigImg = bigImg;
	}
	
	public String getSmallImg() {
		return smallImg;
	}
	
	public void setSmallImg(String smallImg) {
		this.smallImg = smallImg;
	}
	
	public Integer getId() {
		return id;
	}
	public void setId(Integer id) {
		this.id = id;
	}
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	
	public void setDescription(String description) {
		this.description = description;
	}
	
	public String getDescription() {
		return description;
	}
	
	public String getStyle() {
		return style;
	}
	public void setStyle(String style) {
		this.style = style;
	}
	
	
	
	@Override
	public String toString() {
		return "Category [id=" + id + ", name=" + name + ", description="
				+ description + ", bigImg=" + bigImg + ", smallImg=" + smallImg
				+ ", style=" + style + "]";
	}
	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + ((id == null) ? 0 : id.hashCode());
		return result;
	}
	@Override
	public boolean equals(Object obj) {
		if (this == obj)
			return true;
		if (obj == null)
			return false;
		if (getClass() != obj.getClass())
			return false;
		Category other = (Category) obj;
		if (id == null) {
			if (other.id != null)
				return false;
		} else if (!id.equals(other.id))
			return false;
		return true;
	}
	
	
}
