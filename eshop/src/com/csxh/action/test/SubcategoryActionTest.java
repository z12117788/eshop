package com.csxh.action.test;

import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;

import com.csxh.action.SubcategoryAction;

public class SubcategoryActionTest {

	@Before
	public void setUp() throws Exception {
	}

	@Test
	public void testHandle() {
		
		SubcategoryAction action=new SubcategoryAction();
		action.setId(1);
		Assert.assertEquals("success", action.handle());
		
	}

}
