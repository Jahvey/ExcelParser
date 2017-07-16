package com.amosannn.dao;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.amosann.util.DBUtils;
import com.amosannn.entity.UserBean;



public class MysqlToExcel {
	
	public MysqlToExcel() throws SQLException {
		String sql = "select * from user";
	    ResultSet executeQuery = DBUtils.getConnection().createStatement().executeQuery(sql);
		ArrayList<UserBean> beans = new ArrayList<UserBean>();

	    //将数据封装成java bean放到集合中，方便后续使用
	    while(executeQuery.next()){
	        int id = executeQuery.getInt(1);
	        String name = executeQuery.getString(2);
	        String address = executeQuery.getString(3);
	        String email = executeQuery.getString(4);
	        String phone = executeQuery.getString(5);
	        int age = executeQuery.getInt(6);
	        String password = executeQuery.getString(7);
	        beans.add(new UserBean(id, name, address, email, phone, password, age));
	    }
	
	    //开始写入到excel文件中
	
	    //1.创建工作薄
	    XSSFWorkbook book = new XSSFWorkbook();
	
	    //2.创建工作布
	    XSSFSheet sheet = book.createSheet();
	
	    //3.title --- 用户管理列表
	    XSSFCellStyle titleStyle = book.createCellStyle();
	    titleStyle.setAlignment(HorizontalAlignment.CENTER);// 左对齐
	
	    XSSFFont titleFont = book.createFont();
	    titleFont.setBold(true);
	    titleFont.setFontHeightInPoints((short) 18);// 字体大小
	
	    titleStyle.setFont(titleFont);
	
	    CellRangeAddress region = new CellRangeAddress(0, 1, 0, 6);
	    sheet.addMergedRegion(region);
	
	    XSSFCell createCell = sheet.createRow(0).createCell(0);
	    createCell.setCellValue("用户管理列表");
	    createCell.setCellStyle(titleStyle);
	
	
	    //4.设置列描述
	    XSSFRow headerRow = sheet.createRow(2);
	    headerRow.createCell(0).setCellValue("用户ID");
	    headerRow.createCell(1).setCellValue("名称");
	    headerRow.createCell(2).setCellValue("地址");
	    headerRow.createCell(3).setCellValue("邮箱");
	    headerRow.createCell(4).setCellValue("手机号码");
	    headerRow.createCell(5).setCellValue("年龄");
	    headerRow.createCell(6).setCellValue("密码");
	
	
	    for(int i = 0; i < beans.size() ; i++){
	        //5.创建行
	        XSSFRow createRow = sheet.createRow(i+3);
	
	        UserBean bean = beans.get(i);
	
	        //6.设置每列的数据
	        createRow.createCell(0).setCellValue(bean.getId());
	        createRow.createCell(1).setCellValue(bean.getName());
	        createRow.createCell(2).setCellValue(bean.getAddress());
	        createRow.createCell(3).setCellValue(bean.getEmail());
	        createRow.createCell(4).setCellValue(bean.getPhone());
	        createRow.createCell(5).setCellValue(bean.getAge());
	        createRow.createCell(6).setCellValue(bean.getPassword());
	        
	        System.out.println(bean.getId()+" "+bean.getName()+" "+bean.getAddress());
	    }
	
	
	    try {
	        //7.写入，这里测试使用XSSFWorkbook对象写xls以及xlsx可以
	        FileOutputStream fileOutputStream = new FileOutputStream(new File("//Users//jasn//Downloads//test.xlsx"));
	        book.write(fileOutputStream);
	        book.close();
	        fileOutputStream.close();
	        System.out.println("output successful!!!");
	    } catch (IOException e) {
	        e.printStackTrace();
	    }
	
	}
}
