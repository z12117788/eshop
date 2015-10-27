package com.csxh.model;

public class Pager {

	//总的记录数：从数据库查询等到
	private int totalRows;
	
	//每页最多显示的记录数
	private int pageRows;
	
	//共多少页
	private int pageCount;
	
	//当前是第几页，从1开始计算
	private int currentPage=1;
	

	public Pager(int totalRows, int pageSize) {
		super();
		
		this.totalRows = totalRows;
		this.pageRows = pageSize;
		this.pageCount=(int)Math.ceil((double)this.totalRows/(double)this.pageRows);
		
	}
	
	//当前页第一条记录在数据库的编号：以0开始的
	public int getFirstRow(){
		return (this.currentPage-1)*this.pageRows;
	}

	public int getTotalRows() {
		return totalRows;
	}

	public void setTotalRows(int totalRows) {
		this.totalRows = totalRows;
	}

	public int getPageRows() {
		return pageRows;
	}
	
	public void setPageRows(int pageRows) {
		this.pageRows = pageRows;
	}
	
	public int getCurrentPage() {
		return currentPage;
	}

	public void setCurrentPage(int currentPage) {
		this.currentPage = currentPage;
	}

	public int getPageCount() {
		return pageCount;
	}
	
	public boolean hasPrev(){
		return this.currentPage-1>0;
	}
	
	public boolean hasNext(){
		return this.currentPage+1<=this.pageCount;
	}
	
}
