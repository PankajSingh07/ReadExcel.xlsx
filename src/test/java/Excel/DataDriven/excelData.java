package Excel.DataDriven;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class excelData {

	public static void main(String[] args) throws Exception {

		File src = new File("C:\\Users\\pankaj.singh\\Desktop\\Excel Data Driven.xlsx");

		FileInputStream fis = new FileInputStream(src);

		XSSFWorkbook wb = new XSSFWorkbook(fis);

		// For Sheet 1

		XSSFSheet Sheet1 = wb.getSheetAt(0);
//		String data0= sheet1.getRow(0).getCell(0).getStringCellValue();
//		System.out.println("Data of Sheet1 in 1st row and 1st column : "+data0);

		int rowcount1 = Sheet1.getLastRowNum();

		Row r1 = Sheet1.getRow(rowcount1);
		int cellcount1 = r1.getRowNum();

		System.out.println("No. of rows :  " + rowcount1);

		for (int i = 0; i <= cellcount1; i++) {
			System.out.println("--------------------------------------");
			for (int j = 0; j <= rowcount1; j++) 
			{
				String data = Sheet1.getRow(j).getCell(i).getStringCellValue();
				System.out.println("" + data);
			}

		}

		// For Sheet 2
		/*
		  XSSFSheet Sheet2 = wb.getSheetAt(1); // String data1 =
		  Sheet2.getRow(0).getCell(0).getStringCellValue(); //
		  System.out.println("Data from sheet2 of 1st row and 2nd column : " +data1);
		  
		  int rowcount2 = Sheet2.getLastRowNum();
		  
		  Row r2 = Sheet2.getRow(rowcount2); int cellcount2 = r2.getRowNum();
		  
		  System.out.println("No. of rows :  " +rowcount2);
		  
		  for(int i = 0; i<=cellcount2;i++) {
		  System.out.println("--------------------------------------"); 
		 
		  for (int j=0;j<=rowcount2; j++) 
		  {
		  String data =
		  Sheet2.getRow(j).getCell(i).getStringCellValue();
		  System.out.println(""+data); }
		  
		  }
		 */
		wb.close();

	}

}
