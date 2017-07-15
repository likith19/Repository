package com.prc.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class Test1 {

	public static void main(String[] args) {
		FileInputStream in=null;
		FileOutputStream out=null;
		try {
			in = new FileInputStream(new File("C:/Users/likith/Desktop/ExcelPOI/UserData3.xls"));
			out = new FileOutputStream(new File("C:/Users/likith/Desktop/ExcelPOI/UserData3.xls"));
			Workbook workbook1 =new HSSFWorkbook(in);
			Workbook workbook2 =new HSSFWorkbook();
			Sheet firstSheet = workbook1.getSheetAt(0);
			Sheet secondSheet = workbook1.getSheetAt(1);


			for(int k=1;k<4;k++){
				Row row=secondSheet.createRow(k);
				row.createCell(0).setCellValue(firstSheet.getRow(k).getCell(0).getStringCellValue());
				for(int i=1;i<4;i++){
					for(int j=1;j<4;j++){
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
