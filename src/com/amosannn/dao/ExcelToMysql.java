package com.amosannn.dao;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.ArrayList;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.amosann.util.DBUtils;
import com.amosannn.entity.UserBean;

public class ExcelToMysql {
	public ExcelToMysql() throws IOException, SQLException {
		String filePath = "//Users//jasn//Downloads//test.xlsx";

		ArrayList<UserBean> beans = null;
	    FileInputStream input;
	    
		input = new FileInputStream(new File(filePath));

	    Workbook workBook = null;

	    //读取xls和xlsx文件的时候，使用的实例对象并不一样，我们这里进行判断
	    if(filePath.endsWith(".xlsx")){ // 2007版本的excel
	        workBook = new XSSFWorkbook(input);
	    }else{ //2003版的excel
	        workBook = new HSSFWorkbook(input);
	    }

	    beans = new ArrayList<UserBean>();
	    //获取工作布
	    int numberOfSheets = workBook.getNumberOfSheets();
	    for(int i = 0; i < numberOfSheets ; i ++){
	        Sheet sheet = workBook.getSheetAt(i);

	        int lastRowNum = sheet.getLastRowNum();
	        //根据规则，我们从第4行开始读取
	        for(int j = 3 ; j <= lastRowNum ; j ++){
	            Row row = sheet.getRow(j);
	            int id = (int) row.getCell(0).getNumericCellValue();
	            String username = row.getCell(1).getStringCellValue();
	            String address = row.getCell(2).getStringCellValue();
	            String email = row.getCell(3).getStringCellValue();
	            String phone = row.getCell(4).getStringCellValue();
	            int age = (int) row.getCell(5).getNumericCellValue();
	            String pass = row.getCell(6).getStringCellValue();
	            beans.add(new UserBean(id,username,address,email,phone,pass,age));
	        }
	    }


	    input.close();
	    input = null;
	    workBook.close();
	    workBook = null;
		
	    //写入到数据库中
	    Connection connection = DBUtils.getConnection();
	    String sql = "insert into user(username,phone,email,address,age,pass) value(?,?,?,?,?,?)";
	    PreparedStatement prepareStatement = connection.prepareStatement(sql);
	    for(int i = 0 ;i < beans.size(); i ++){
	        UserBean bean = beans.get(i);
	        prepareStatement.setString(1, bean.getName());
	        prepareStatement.setString(2, bean.getPhone());
	        prepareStatement.setString(3, bean.getEmail());
	        prepareStatement.setString(4, bean.getAddress());
	        prepareStatement.setInt(5, bean.getAge());
	        prepareStatement.setString(6, bean.getPassword());
	        //添加到批处理中
	        prepareStatement.addBatch();
	    }

	    //执行批处理
	    int[] executeBatch = prepareStatement.executeBatch();
	    System.out.println(executeBatch);
	}
}
