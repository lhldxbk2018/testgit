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
	        //��һ��������һ��workbook��Ӧһ��Excel�ļ�
	        HSSFWorkbook workbook=new HSSFWorkbook();
	        //�ڶ�������workbook�д���һ��sheet��ӦExcel�е�sheet
	        HSSFSheet sheet=workbook.createSheet("�û���һ");
	        //����������sheet������ӱ�ͷ��0�У��ϰ汾��poi��sheet������������
	        HSSFRow row=sheet.createRow(0);
	        //���Ĳ���������Ԫ�����ñ�ͷ
	        HSSFCell cell=row.createCell((short) 0);
	        cell.setCellValue("�û���");
	        cell=row.createCell((short) 1);
	        cell.setCellValue("����");

	        //���岿��д��ʵ�����ݣ�ʵ��Ӧ������Щ ���ݴ����ݿ�õ��������װ���ݣ����ϰ����󡣶��������ֵ��Ӧ���ÿ�е�ֵ
	        for(int i=0;i<list.size();i++){
	            HSSFRow row1=sheet.createRow(i+1);
	            User user=list.get(i);
	            //������Ԫ����ֵ
	            row1.createCell((short)0).setCellValue(user.getUserAccount());
	            row1.createCell((short)1).setCellValue(user.getPassword());
	        }
	        //���ļ����浽ָ����λ��
	        try {
	            FileOutputStream fos=new FileOutputStream("E:\\user.xls");
	            workbook.write(fos);
	            System.out.println("д��ɹ�");
	            fos.close();
	        } catch (IOException e) {
	            // TODO Auto-generated catch block
	            e.printStackTrace();
	        }

	    }
	  

	}


