package com.amosannn.run;

import java.io.IOException;
import java.sql.SQLException;

import com.amosannn.dao.ExcelToMysql;
import com.amosannn.dao.MysqlToExcel;

public class run {
	public static void main(String args[]){
		
	    try {
			new MysqlToExcel();
//	    	new ExcelToMysql();
		} catch (SQLException e) {
			e.printStackTrace();
		}

	}
}
