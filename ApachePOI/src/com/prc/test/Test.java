package com.prc.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class Test {

	public static void main(String[] args) {
		FileInputStream in=null;
		FileOutputStream out=null;
		try {
			 in = new FileInputStream(new File("C:/Users/likith/Desktop/ExcelPOI/UserData1.xls"));
			 out = new FileOutputStream(new File("C:/Users/likith/Desktop/ExcelPOI/UserData2.xls"));
			Workbook workbook1 =new HSSFWorkbook(in);
			Sheet firstSheet = workbook1.getSheetAt(0);
			Workbook workbook2 =new HSSFWorkbook();
			Sheet secondSheet = workbook2.createSheet("Group Details");

			for(int k=1;k<4;k++){//No of rows
				Row row=secondSheet.createRow(k);
				row.createCell(0).setCellValue(firstSheet.getRow(k).getCell(0).getStringCellValue());
				for(int i=1;i<4;i++){//No of columns
						for(int j=1;j<4;j++){//No of rows
							if(null != firstSheet.getRow(j).getCell(i) && 
									(firstSheet.getRow(k).getCell(0).getStringCellValue().equalsIgnoreCase(firstSheet.getRow(j).getCell(i).getStringCellValue()))){
								System.out.println(firstSheet.getRow(j).getCell(i).getStringCellValue());
								System.out.println("K:" +k +"i:"+i+"j:"+j);
								row.createCell(i).setCellValue((String)firstSheet.getRow(j).getCell(i).getStringCellValue());
							}	
					}
				}
			}
			
			workbook2.write(out);
			
		} catch (Exception e) {
			e.printStackTrace();
		}finally{
			try {
				out.flush();
				out.close();
				in.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

}
