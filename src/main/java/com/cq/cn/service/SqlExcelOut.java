package com.cq.cn.service;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import com.cq.cn.bean.User;

@WebServlet("/export")
public class SqlExcelOut extends HttpServlet {
	
	private static final long serialVersionUID = 1L;
	
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		//第一步，创建一个webbook，对应一个Excel文件
		HSSFWorkbook workbook = new HSSFWorkbook();
		//第二步，在webbook中添加一个sheet,对应Excel文件中的sheet
		HSSFSheet sheet = workbook.createSheet("first");
		//第三步，在sheet中添加表头第0行,注意老版本poi对Excel的行数列数有限制short
		HSSFRow row = sheet.createRow((int)0);
		//第四步，创建单元格，并设置值表头 设置表头居中
		HSSFCellStyle style = workbook.createCellStyle();
		style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		HSSFCell cell0 = row.createCell((short) 0);
		cell0.setCellValue("姓名");
		cell0.setCellStyle(style);
		HSSFCell cell1 = row.createCell((short) 1);
		cell1.setCellValue("年龄");
		cell1.setCellStyle(style);
		//第五步，写入实体数据 实际应用中这些数据从数据库得到
		User user1 = new User();
		user1.setUname("李雷");
		user1.setAge(12);
		User user2 = new User();
		user2.setUname("张三");
		user2.setAge(22);
		List<User> list = new ArrayList<User>();
		list.add(user1);
		list.add(user2);
		for (int i = 0; i < list.size(); i++) {
			HSSFRow row2 = sheet.createRow((int)i+1);
			HSSFCell cell8 = row2.createCell((short) 0);
			HSSFCell cell9 = row2.createCell((short) 1);
			cell8.setCellStyle(style);
			cell9.setCellStyle(style);
			cell8.setCellValue(list.get(i).getUname());
			cell9.setCellValue(list.get(i).getAge());
		}
		//第六步，将文件存到指定位置
		FileOutputStream out = new FileOutputStream("E:\\myexcel.xls");
		workbook.write(out);
		response.getWriter().write("ok");
		out.close();
	}
	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		doGet(request, response);
	}
}