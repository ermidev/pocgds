package com.expedia.gds;

import java.io.*;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.*;

public class Program {
	private static XSSFWorkbook workbook;
	private static XSSFSheet spreadSheet;

	public static void main(String[] args) throws Exception{
	      workbook = new XSSFWorkbook(); 
	      spreadSheet = workbook.createSheet("First");
	      
	      XSSFRow row = spreadSheet.createRow((short)1);
	      
	      Map<String, Object[]> rowValues = new TreeMap<String, Object[]>();
	      
	      // write database access object here.
	      
	      rowValues.put("1", new Object[]{"FirstName", "LastName", "Department"});
	      rowValues.put("2", new Object[]{"Ermias", "Kidane", "IT"});
	      rowValues.put("3", new Object[]{"Ermias", "Kidane", "IT"});
	      rowValues.put("4", new Object[]{"Ermias", "Kidane", "IT"});
	      rowValues.put("5", new Object[]{"Ermias", "Kidane", "IT"});
	      rowValues.put("6", new Object[]{"Ermias", "Kidane", "IT"});
	      rowValues.put("7", new Object[]{"Ermias", "Kidane", "IT"});
	      rowValues.put("8", new Object[]{"Ermias", "Kidane", "IT"});
	      rowValues.put("9", new Object[]{"Cathy", "Xu", "IT"});
	      
	      Set<String> keys = rowValues.keySet();
	      int rowId = 0;
	      for(String key : keys){
	    	  row = spreadSheet.createRow(rowId++);
	    	  
	    	  Object[] objectArr = rowValues.get(key);
	    	  int cellId = 0;
	    	  for(Object obj: objectArr){
	    		  Cell cell = row.createCell(cellId++);
	    		  cell.setCellValue((String)obj);
	    	  }
	      }
	      //Create file system using specific name
	      FileOutputStream out = new FileOutputStream(
	      new File("createworkbook.xlsx"));
	      //write operation workbook using file out object 
	      workbook.write(out);
	      out.close();
	      System.out.println("createworkbook.xlsx written successfully");
	      
		System.out.println("test");
	}
}
