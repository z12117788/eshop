package com.csxh.util;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;

public class JdbcUtil {

	static String driver = "com.microsoft.sqlserver.jdbc.SQLServerDriver";
	static String url = "jdbc:sqlserver://localhost:1433;DatabaseName=db";
	static String user = "sa";
	static String password = "1234";

	// 使用静态块初始化静态变量
	static {

		try {
			// 从属性配置文件中读取参数
			Properties p = new Properties();
			// 从类的路径（classpath）读取文件，方法是通过类加载器来获取
			InputStream is = JdbcUtil.class.getClassLoader()
					.getResourceAsStream("jdbc.properties");
			p.load(is);
			JdbcUtil.driver = p.getProperty("driver");
			JdbcUtil.url = p.getProperty("url");
			JdbcUtil.user = p.getProperty("user");
			JdbcUtil.password = p.getProperty("password");
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	// 打开数据库连接
	public static Connection openConnection() throws SQLException,
			ClassNotFoundException {

		Connection cnn = null;

		// 通过反射加载数据库驱动类的对象
		Class.forName(JdbcUtil.driver);
		// 就可以获取连接对象
		cnn = DriverManager.getConnection(JdbcUtil.url, JdbcUtil.user,
				JdbcUtil.password);

		return cnn;

	}

	public static int queryTotalRows(String tableName, String pramaryKey) {
		int totalRows = 0;
		String sql = "select count(" + pramaryKey + ") from " + tableName;
		try {
			Connection cnn = openConnection();
			Statement stmt = cnn.createStatement();
			ResultSet rs = stmt.executeQuery(sql);
			if (rs.next()) {
				totalRows = rs.getInt(1);
			}
		} catch (ClassNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return totalRows;
	}

	public static int queryTotalRows(String tableName, String pramaryKey,String where) {
		int totalRows = 0;
		where=where.replace("where", "");
		String sql = "select count(" + pramaryKey + ") from " + tableName +" where "+where;
		try {
			Connection cnn = openConnection();
			Statement stmt = cnn.createStatement();
			ResultSet rs = stmt.executeQuery(sql);
			if (rs.next()) {
				totalRows = rs.getInt(1);
			}
		} catch (ClassNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return totalRows;
	}

	// 查询数据库中表的记录
	public static List<Object[]> queryForObjectList(String sql) {

		List<Object[]> objectList = new ArrayList<Object[]>();
		// 执行SQL语句获取结果
		ResultSet rs = null;
		Connection cnn = null;
		Statement stmt = null;
		try {
			// 打开连接对象
			cnn = JdbcUtil.openConnection();
			// 创建SQL语句对象
			stmt = cnn.createStatement();
			// 执行SQL语句
			rs = stmt.executeQuery(sql);

			// 在rs关闭之前，将其中的内容放入到objectList中
			// 首先判断一行有几列数据
			ResultSetMetaData rsmd = rs.getMetaData();

			int columnCount = rsmd.getColumnCount();

			while (rs.next()) {// 每一行的数据，一行中就有列的数据

				Object[] objects = new Object[columnCount];
				// 每一项对应就是列的数据
				for (int column = 0; column < columnCount; column++) {
					objects[column] = rs.getObject(column + 1);// 将第column+1列的数据放入到数据的第column项
				}

				objectList.add(objects);

			}

			// 关闭数据库资源:rs,stmt,cnn
			// 为了确保所的资源都关闭，必须将关闭资源的代码放在finally块执行
			// rs.close();
			// stmt.close();
			// cnn.close();

		} catch (ClassNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {// 该块能够在返回时提前执行，无论是否发生异常，都会执行该块中的代码
			try {
				if (rs != null) {
					rs.close();
				}
				if (stmt != null) {
					stmt.close();
				}
				if (cnn != null) {
					cnn.close();
				}
			} catch (SQLException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}

		// 返回结果集
		return objectList;

	}

	// 向数据库表中插入记录：带参数的SQL需要使用占位符类似于format中占位符，但占位符统一用'?'表示
	public static boolean insertObject(String sql, Object... values) {// insert
																		// into
																		// 用户
																		// values(?,?,?,?);

		boolean fOk = false;

		Connection cnn = null;
		// 使用带参数的命令语句
		PreparedStatement pstmt = null;

		try {
			// 打开连接对象
			cnn = JdbcUtil.openConnection();
			// 准备命令语句对象
			pstmt = cnn.prepareStatement(sql);
			// 设置占位符对应的值（要考虑值的类型）
			for (int i = 0; i < values.length; i++) {
				// Map<Object,Object> map=values[i];
				// //取第i个占位符的类型
				// Object type=map.keySet().iterator().next();//只取第一个条目的key
				// //取第i个占位符的类型对值的值
				// Object value=map.get(type);

				// 根据类型来调用对应的方法
				if (values[i] instanceof Integer) {
					pstmt.setInt(i + 1, (Integer) values[i]);// 已经知道是Integer类型
				} else if (values[i] instanceof String) {
					pstmt.setString(i + 1, (String) values[i]);
				} else if (values[i].getClass().getName()
						.equals("java.lang.Boolean")) {
					pstmt.setBoolean(i + 1, (Boolean) values[i]);
				} else {
					// throw new Exception("类型不认识");
				}

			}

			// 设置好各占位符的值后，就执行SQL语句
			int rs = pstmt.executeUpdate();

			fOk = rs > 0;// 插入成功的条件是受影响的行数要大于0

		} catch (ClassNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			try {

				if (pstmt != null) {
					pstmt.close();
				}

				if (cnn != null) {
					cnn.close();
				}
			} catch (SQLException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}

		return fOk;

	}

	// 向数据库表插入记录--->JavaBean对象（所有JavaBean对象都是Object对象）
	// 前题条件：1、JavaBean的属性名称（getXxxx）--->表的列的名称要一致,2、JavaBean类的名称---->表的名称
	// User username---->表username列
	public static boolean insertObject(Object javaBean, String... excludeFields) {// User---user

		boolean fOk = false;

		// 要动态生成sql语句
		StringBuilder sb = new StringBuilder();
		sb.append("insert into ");
		// 通过对象的返反射对象获取类的名称（简短名称，不带包的名称）
		String simpleName = javaBean.getClass().getSimpleName();
		char firstChar = simpleName.charAt(0);
		firstChar = java.lang.Character.toLowerCase(firstChar);
		// 构建表的名称
		sb.append('[').append(firstChar).append(simpleName.substring(1))
				.append("]");

		// 构建表的列名称
		sb.append(" ( ");

		// 保存所有的getXxxx方法，用于取值
		List<Method> getMethodList = new ArrayList<Method>();
		// 通过对象的反射对象获取每一个属性的名称
		Method[] methods = javaBean.getClass().getMethods();
		for (int i = 0; i < methods.length; i++) {
			String mn = methods[i].getName();// 获取第i个方法的名称
			// 判断是否是get方法
			if (mn.startsWith("get")) {
				// 将第4个字符变成小写
				firstChar = mn.charAt(3);
				firstChar = java.lang.Character.toLowerCase(firstChar);

				String field = mn.substring(4);
				for (String exclude : excludeFields) {
					if ((firstChar + field).equals(exclude)) {
						field = null;// 如果该字段是被排除就设为null
						break;
					}
				}

				if (field != null) {// 如果不为null，则是要插入的列名

					// 判断该方法是否有值
					Object value = null;
					try {
						value = methods[i].invoke(javaBean);
					} catch (IllegalAccessException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					} catch (IllegalArgumentException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					} catch (InvocationTargetException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
					if (null != value) {
						// 说明该值对象的列值需要更新

						// 拼接成：列名=？,的形式
						sb.append(firstChar);// 拼接成属性名称第一个字符
						sb.append(field);// 拼接成属性名称后面的字符
						sb.append("=?,");

						// 保存这个get方法的反射对象
						getMethodList.add(methods[i]);
					}

				}

			}// end if(mn.startsWith("get"))

		}// end for

		// 构建where语句

		// 将最后一个逗号删除
		sb.deleteCharAt(sb.length() - 1);

		sb.append(" )  ");

		// 最后构建值的列表
		sb.append("values (");

		// 使用占位符来代替对应的值
		for (int i = 0; i < getMethodList.size(); i++) {
			sb.append("?,");
		}

		sb.deleteCharAt(sb.length() - 1);

		sb.append(")");

		System.out.println(sb.toString());

		// 使用连接对象创建带参数的命令语句对象
		Connection cnn = null;
		PreparedStatement ps = null;

		String sql = sb.toString();
		try {

			cnn = JdbcUtil.openConnection();

			ps = cnn.prepareStatement(sql);

			// 为每一个占位符设置对应的值（要考虑值的类型）
			for (int i = 0; i < getMethodList.size(); i++) {
				Method m = getMethodList.get(i);// 获取占位符对应get方法
				// 通过反射方法来调用get方法获取对应的值
				Object value = m.invoke(javaBean);
				// if(value instanceof String){
				// ps.setString(i+1, (String)value);
				// }else if (value instanceof Integer){
				// ps.setInt(i+1, (Integer)value);
				// }
				ps.setObject(i + 1, value);// 设置占位符的值
			}
			// fOk=ps.execute();
			int rc = ps.executeUpdate();

			fOk = rc > 0;
		} catch (ClassNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IllegalAccessException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IllegalArgumentException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (InvocationTargetException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			try {

				if (ps != null) {
					ps.close();
				}
				if (cnn != null) {
					cnn.close();
				}

			} catch (SQLException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}

		return fOk;

	}

	// 向数据库表中更新记录
	// 如果JavaBean对象的get方法返回的值不为空，则该属性对应的列需要更新
	public static boolean updateObject(Object javaBean, String id) {

		boolean fOk = false;

		// 动态生成update语句
		StringBuilder sb = new StringBuilder();
		sb.append("update  ");
		// 通过对象的返反射对象获取类的名称（简短名称，不带包的名称）
		String simpleName = javaBean.getClass().getSimpleName();
		char firstChar = simpleName.charAt(0);
		firstChar = java.lang.Character.toLowerCase(firstChar);
		// 构建表的名称
		sb.append('[').append(firstChar).append(simpleName.substring(1))
				.append("]");
		// 构建set语句
		sb.append("  set ");

		// 设置条件的get方法
		Map<String, Object> whereField2ValueMap = new HashMap<String, Object>();

		// 保存所有的列名与列值映射列表
		List<Map<String, Object>> updateField2ValueMapList = new ArrayList<Map<String, Object>>();
		// 通过对象的反射对象获取每一个属性的名称
		Method[] methods = javaBean.getClass().getMethods();
		for (int i = 0; i < methods.length; i++) {
			String mn = methods[i].getName();// 获取第i个方法的名称
			// 判断是否是get方法
			if (mn.startsWith("get")) {
				// 将第4个字符变成小写
				firstChar = mn.charAt(3);
				firstChar = java.lang.Character.toLowerCase(firstChar);

				// 获取列名
				String field = firstChar + mn.substring(4);
				if (field.equals("class")) {
					continue;// 立即进入下一次循环，不再执行后的语句了
				}
				// 获取列值
				Object value = null;
				try {
					value = methods[i].invoke(javaBean);
				} catch (IllegalAccessException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (IllegalArgumentException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (InvocationTargetException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}

				if (value != null) {// 如果不为null，则是要插入的列名

					// 判断该列表是否为条件列名
					if (field.equals(id)) {// 说明是一个条件列-值对
						whereField2ValueMap.put(field, value);// 保存下来，为构建where语句作准备
					} else {
						// 说明是要更新的列-值对象
						sb.append(field);// 设置列名
						sb.append("=?,");

						// 创建列名-列值映射对象
						Map<String, Object> field2ValueMap = new HashMap<>();
						field2ValueMap.put(field, value);
						// 将列名-列值映射对象加入列表中
						updateField2ValueMapList.add(field2ValueMap);
					}

				}

			}// end if(mn.startsWith("get"))

		}// end for

		// 删除最后的逗号
		sb.deleteCharAt(sb.length() - 1);

		// 构建where语句
		sb.append(" where ").append(
				whereField2ValueMap.keySet().iterator().next());
		sb.append("=?");

		System.out.println(sb);

		// 使用连接与语句对象
		Connection cnn = null;
		PreparedStatement ps = null;

		try {
			cnn = JdbcUtil.openConnection();
			ps = cnn.prepareStatement(sb.toString());

			// 设置占位符的值
			int index = 1;
			for (int i = index - 1; i < updateField2ValueMapList.size(); i++, index++) {
				Map<String, Object> field2ValueMap = updateField2ValueMapList
						.get(i);
				ps.setObject(index, field2ValueMap.values().iterator().next());
			}

			ps.setObject(index, whereField2ValueMap.values().iterator().next());

			int rc = ps.executeUpdate();
			fOk = rc > 0;
		} catch (SQLException | ClassNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			try {

				if (ps != null) {
					ps.close();
				}

				if (cnn != null) {
					cnn.close();
				}
			} catch (SQLException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}

		return fOk;

	}

	// 从数据库表中删除记录
	public static boolean deleteObject(Object javaBean, String key) {// key=id
																		// --getId,key=uid--getUid
		boolean fOk = false;

		StringBuilder sb = new StringBuilder();
		sb.append("delete from ");
		String simpleName = javaBean.getClass().getSimpleName();
		char firstChar = simpleName.charAt(0);
		firstChar = java.lang.Character.toLowerCase(firstChar);
		sb.append("[");
		sb.append(firstChar).append(simpleName.substring(1));
		sb.append("]");

		// 寻找key对应get方法,并获取其值
		Object value = null;
		Method[] methods = javaBean.getClass().getDeclaredMethods();
		for (Method method : methods) {
			String getMethod = "get"
					+ java.lang.Character.toUpperCase(key.charAt(0))
					+ key.substring(1);
			if (getMethod.equals(method.getName())) {
				try {
					value = method.invoke(javaBean);
				} catch (IllegalAccessException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (IllegalArgumentException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (InvocationTargetException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				break;
			}
		}

		if (value != null) {

			// 不使用占位符
			if (value instanceof String) {
				sb.append(" where ").append(key)
						.append("='" + value.toString() + "' ");// 因为如果是值是字符串，则要加上一对单引号
			} else {
				sb.append(" where ").append(key)
						.append("=" + value.toString() + " ");// 因为如果不是值是字符串，则不要加上一对单引号
			}

		}

		System.out.println(sb);

		Connection cnn = null;
		Statement stmt = null;

		try {
			cnn = JdbcUtil.openConnection();
			stmt = cnn.createStatement();
			int n = stmt.executeUpdate(sb.toString());
			fOk = n > 0;
		} catch (ClassNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {

			try {
				if (stmt != null) {
					stmt.close();
				}
				if (cnn != null) {
					cnn.close();
				}
			} catch (SQLException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}

		}

		return fOk;
	}

	// 从数据库表中查询记录
	public static <T> List<T> queryObjectList(Class<T> javaBeanClass,
			String... wheres) {

		List<T> objectList = new ArrayList<T>();

		// 动态构建SQL语句
		StringBuilder sb = new StringBuilder();
		sb.append("select * from ");
		String simpleName = javaBeanClass.getSimpleName();
		char firstChar = simpleName.charAt(0);
		firstChar = java.lang.Character.toLowerCase(firstChar);
		sb.append("[");
		sb.append(firstChar).append(simpleName.substring(1));
		sb.append("]");

		// 判断是否有条件语句
		if (wheres.length > 0) {

			sb.append(" where  ");
			for (String where : wheres) {
				sb.append(where).append(" and ");
			}

			sb.delete(sb.length() - 5, sb.length());

		}

		System.out.println(sb);

		// 进行查询操作
		Connection cnn = null;
		Statement stmt = null;
		ResultSet rs = null;

		try {
			cnn = JdbcUtil.openConnection();
			stmt = cnn.createStatement();
			rs = stmt.executeQuery(sb.toString());
			// 通过结果集对象获取关于查询的元数据对象
			ResultSetMetaData rsmd = rs.getMetaData();
			int columnCount = rsmd.getColumnCount();

			while (rs.next()) {

				// 创建对象有三种方法：1、new 2、class.forName 3、class.newInstance
				T object = null;
				try {
					object = javaBeanClass.newInstance();
				} catch (InstantiationException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (IllegalAccessException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				// 每一条记录---->一个javabean对象
				// 列的名称=javabean.属性(setXxxx);
				for (int i = 0; i < columnCount; i++) {
					// 获取列的名称
					String columnName = rsmd.getColumnName(i + 1);
					// 获取列的值
					Object value = rs.getObject(i + 1);
					// 获取列的值的类型的名称
					String columnValueClass = rsmd.getColumnClassName(i + 1);

					// 根据列的名称来构建JavaBean对象的setXxxx方法
					String setXxxx = "set"
							+ java.lang.Character.toUpperCase(columnName
									.charAt(0)) + columnName.substring(1);

					Method[] methods = javaBeanClass.getMethods();
					// 寻找setXxxx方法，并调用它
					for (Method method : methods) {
						if (method.getName().equals(setXxxx)) {
							try {
								method.invoke(object, value);// 实际上是调用JavaBean对象的setXxxx方法
							} catch (IllegalAccessException e) {
								// TODO Auto-generated catch block
								e.printStackTrace();
							} catch (IllegalArgumentException e) {
								// TODO Auto-generated catch block
								e.printStackTrace();
							} catch (InvocationTargetException e) {
								// TODO Auto-generated catch block
								e.printStackTrace();
							}

							break;
						}
					}// end for method

				}// end for column

				// 将JavaBean对象加入返回列表中
				objectList.add(object);

			}
		} catch (ClassNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			try {
				if (rs != null) {
					rs.close();
				}
				if (stmt != null) {
					stmt.close();
				}
				if (cnn != null) {
					cnn.close();
				}
			} catch (SQLException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}

		return objectList;
	}

	// 当表与javabean的名称不一致，可以建立它们之间的名称映射对象
	/*
	 * Customer---->雇员 id------->雇员ID
	 */
	public static boolean insertObject(Object javaBean,
			Map<String, String> javaBean2TableNameMap, String... excludeFileds) {

		boolean fOk = false;

		// 要动态生成sql语句
		StringBuilder sb = new StringBuilder();
		sb.append("insert into ");
		// 通过对象的返反射对象获取类的名称（简短名称，不带包的名称）
		String simpleName = javaBean.getClass().getSimpleName();
		char firstChar = simpleName.charAt(0);
		firstChar = java.lang.Character.toLowerCase(firstChar);
		String javaBeanName = firstChar + simpleName.substring(1);
		// 构建表的名称
		// 首先查询映射名称
		String tableName = javaBean2TableNameMap.get(javaBeanName);
		if (tableName == null) {
			// 说明表的名称与JavaBean的名称一致
			tableName = javaBeanName;
		}
		sb.append('[').append(tableName).append("]");

		// 构建表的列名称
		sb.append(" ( ");

		// 保存所有的getXxxx方法，用于取值
		List<Method> getMethodList = new ArrayList<Method>();
		// 通过对象的反射对象获取每一个属性的名称
		Method[] methods = javaBean.getClass().getMethods();
		for (int i = 0; i < methods.length; i++) {
			String mn = methods[i].getName();// 获取第i个方法的名称
			// 判断是否是get方法
			if (mn.startsWith("get")) {
				// 将第4个字符变成小写
				firstChar = mn.charAt(3);
				firstChar = java.lang.Character.toLowerCase(firstChar);

				String field = mn.substring(4);
				for (String exclude : excludeFileds) {
					if ((firstChar + field).equals(exclude)) {
						field = null;// 如果该字段是被排除就设为null
						break;
					}
				}

				if (field != null) {// 如果不为null，则是要插入的列名

					// 判断该方法是否有值
					Object value = null;
					try {
						value = methods[i].invoke(javaBean);
					} catch (IllegalAccessException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					} catch (IllegalArgumentException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					} catch (InvocationTargetException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
					if (null != value) {
						// 说明该值对象的列值需要更新

						// 拼接成：列名=？,的形式
						String javaBeanFieldName = firstChar + field;// 拼接成属性名称
						String tableFieldName = javaBean2TableNameMap
								.get(javaBeanFieldName);
						if (tableFieldName == null) {
							// 说明一致
							tableFieldName = javaBeanFieldName;
						}

						sb.append(tableFieldName);//
						sb.append(",");

						// 保存这个get方法的反射对象
						getMethodList.add(methods[i]);
					}

				}

			}// end if(mn.startsWith("get"))

		}// end for

		// 构建where语句

		// 将最后一个逗号删除
		sb.deleteCharAt(sb.length() - 1);

		sb.append(" )  ");

		// 最后构建值的列表
		sb.append("values (");

		// 使用占位符来代替对应的值
		for (int i = 0; i < getMethodList.size(); i++) {
			sb.append("?,");
		}

		sb.deleteCharAt(sb.length() - 1);

		sb.append(")");

		System.out.println(sb.toString());

		// 使用连接对象创建带参数的命令语句对象
		Connection cnn = null;
		PreparedStatement ps = null;

		String sql = sb.toString();
		try {

			cnn = JdbcUtil.openConnection();

			ps = cnn.prepareStatement(sql);

			// 为每一个占位符设置对应的值（要考虑值的类型）
			for (int i = 0; i < getMethodList.size(); i++) {
				Method m = getMethodList.get(i);// 获取占位符对应get方法
				// 通过反射方法来调用get方法获取对应的值
				Object value = m.invoke(javaBean);
				ps.setObject(i + 1, value);// 设置占位符的值
			}
			// fOk=ps.execute();
			int rc = ps.executeUpdate();

			fOk = rc > 0;
		} catch (ClassNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IllegalAccessException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IllegalArgumentException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (InvocationTargetException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			try {

				if (ps != null) {
					ps.close();
				}
				if (cnn != null) {
					cnn.close();
				}

			} catch (SQLException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}

		return fOk;

	}

}
