package com.cq.cn.service;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import com.cq.cn.bean.User;

public class SqlExcel {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		 User user=new User();
	        List<User> list=new ArrayList<User>();
	        user.setUserAccount("admin");
	        user.setPassword("admin");
	        list.add(user);
	        User user1=new User();
	        user1.setUserAccount("commonuser");
	        user1.setPassword("123456");
	        list.add(user1);
	        getExcel(list);
	}
	
	
	    public static void getExcel(List<User> list){
	        //第一步：创建一个workbook对应一个Excel文件
	        HSSFWorkbook workbook=new HSSFWorkbook();
	        //第二部：在workbook中创建一个sheet对应Excel中的sheet
	        HSSFSheet sheet=workbook.createSheet("用户表一");
	        //第三部：在sheet表中添加表头第0行，老版本的poi对sheet的行列有限制
	        HSSFRow row=sheet.createRow(0);
	        //第四部：创建单元格，设置表头
	        HSSFCell cell=row.createCell((short) 0);
	        cell.setCellValue("用户名");
	        cell=row.createCell((short) 1);
	        cell.setCellValue("密码");

	        //第五部：写入实体数据，实际应用中这些 数据从数据库得到，对象封装数据，集合包对象。对象的属性值对应表的每行的值
	        for(int i=0;i<list.size();i++){
	            HSSFRow row1=sheet.createRow(i+1);
	            User user=list.get(i);
	            //创建单元格设值
	            row1.createCell((short)0).setCellValue(user.getUserAccount());
	            row1.createCell((short)1).setCellValue(user.getPassword());
	        }
	        //将文件保存到指定的位置
	        try {
	            FileOutputStream fos=new FileOutputStream("E:\\user.xls");
	            workbook.write(fos);
	            System.out.println("写入成功");
	            fos.close();
	        } catch (IOException e) {
	            // TODO Auto-generated catch block
	            e.printStackTrace();
	        }

	    }
	  

	}


