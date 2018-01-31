package com.cq.cn.service;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.fileupload.FileItem;
import org.apache.commons.fileupload.disk.DiskFileItemFactory;
import org.apache.commons.fileupload.servlet.ServletFileUpload;
import org.apache.poi.hssf.model.Sheet;
import org.apache.poi.hssf.record.formula.functions.Cell;
import org.apache.poi.hssf.record.formula.functions.Row;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import com.cq.cn.bean.User;

@WebServlet("/upload")
public class SqlExcelIn extends HttpServlet{
	private static final long serialVersionUID = 1L;

	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		DiskFileItemFactory factory = new DiskFileItemFactory();
		ServletFileUpload upload = new ServletFileUpload(factory);
		try {
			  List<FileItem> items = upload.parseRequest(request);
			  for (FileItem fileItem : items) {
				InputStream in = fileItem.getInputStream();
				HSSFWorkbook workbook = new HSSFWorkbook(in); 
				//HSSFSheet sheet = workbook.getSheet("Sheet1");
				List<User> list = new ArrayList<User>();
				/*for (HSSFRow row :sheet) {
					int rowNum = row.getRowNum();
					if(rowNum>0){
						User user = new sUser();
						String uname = row.getCell(0).getStringCellValue();
						//row.getCell(1).setCellType(Cell.CEsLL_TYPE_STRING);
						String age = row.getCell(1).getStringCellValue();
						int sage = Integer.parseInt(age);
						user.setUname(uname);
						user.setAge(sage);
						list.add(user);
					}
						
					}*/
				
				
			          
			         // 循环工作表Sheet  
			         for (int numSheet = 0; numSheet < workbook.getNumberOfSheets(); numSheet++) {  
			             HSSFSheet hssfSheet = workbook.getSheetAt(numSheet);  
			             if (hssfSheet == null) {  
			                 continue;  
			             }  
			             // 循环行Row  
			             for (int rowNum = 1; rowNum <= hssfSheet.getLastRowNum(); rowNum++) {  
			                 HSSFRow hssfRow = hssfSheet.getRow(rowNum);  
			                 if (hssfRow != null) {  
			                	 User user = new User(); 
			                     String uname = hssfRow.getCell(0).getStringCellValue();  
			                     String age = hssfRow.getCell(1).getStringCellValue(); 
			                     int sage = Integer.parseInt(age);
			                     user.setUname(uname);
								 user.setAge(sage);
								 list.add(user); 
			                 }  
			             }  
			         }  
			         
				
				
				
				
				
				
				
				
				
					request.getSession().setAttribute("list", list);
					response.getWriter().write("ok");
				}
				
					
				
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		doGet(request,response);
	}

	

}


