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
		//��һ��������һ��webbook����Ӧһ��Excel�ļ�
		HSSFWorkbook workbook = new HSSFWorkbook();
		//�ڶ�������webbook�����һ��sheet,��ӦExcel�ļ��е�sheet
		HSSFSheet sheet = workbook.createSheet("first");
		//����������sheet����ӱ�ͷ��0��,ע���ϰ汾poi��Excel����������������short
		HSSFRow row = sheet.createRow((int)0);
		//���Ĳ���������Ԫ�񣬲�����ֵ��ͷ ���ñ�ͷ����
		HSSFCellStyle style = workbook.createCellStyle();
		style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		HSSFCell cell0 = row.createCell((short) 0);
		cell0.setCellValue("����");
		cell0.setCellStyle(style);
		HSSFCell cell1 = row.createCell((short) 1);
		cell1.setCellValue("����");
		cell1.setCellStyle(style);
		//���岽��д��ʵ������ ʵ��Ӧ������Щ���ݴ����ݿ�õ�
		User user1 = new User();
		user1.setUname("����");
		user1.setAge(12);
		User user2 = new User();
		user2.setUname("����");
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
		//�����������ļ��浽ָ��λ��
		FileOutputStream out = new FileOutputStream("E:\\myexcel.xls");
		workbook.write(out);
		response.getWriter().write("ok");
		out.close();
	}
	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		doGet(request, response);
	}
}