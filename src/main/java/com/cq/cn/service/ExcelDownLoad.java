package com.cq.cn.service;


import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URLEncoder;

import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

/**
 * Servlet implementation class DownloadServlet
 */
@WebServlet("/download")
public class ExcelDownLoad extends HttpServlet {
	private static final long serialVersionUID = 1L;
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		try {
            //��ȡ�ļ���·��
            String excelPath = request.getSession().getServletContext().getRealPath("/excel/"+"abc.xlsx");
            String fileName = "abc.xlsx".toString(); // �ļ���Ĭ�ϱ�����
            // ��������
            InputStream inStream = new FileInputStream(excelPath);//�ļ��Ĵ��·��
            // ��������ĸ�ʽ
            response.reset();
            response.setContentType("bin");
            response.addHeader("Content-Disposition",
                    "attachment;filename=" + URLEncoder.encode("abc.xlsx", "UTF-8"));
            // ѭ��ȡ�����е�����
            byte[] b = new byte[200];
            int len;

            while ((len = inStream.read(b)) > 0){
                response.getOutputStream().write(b, 0, len);
            }
            inStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
	}

	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		doGet(request,response);
	}

}