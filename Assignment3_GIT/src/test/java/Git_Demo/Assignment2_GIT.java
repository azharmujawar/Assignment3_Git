package Git_Demo;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;


import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Assignment2_GIT {

	public static void main(String[] args) throws IOException {
		
		
		FileInputStream fis=new FileInputStream(new File("D:\\EmployeeData.xlsx"));
		
		XSSFWorkbook workbook=new XSSFWorkbook(fis);



		XSSFSheet sheet=workbook.getSheetAt(0);

  
  int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();
  
  for(int i=1;i<rowCount;i++) 
  { for(int j=1;j<rowCount;j++)
  { 
	  if(j==1) 
	  {
	  
	  
	  
  System.out.print((int)sheet.getRow(i).getCell(j).getNumericCellValue());
  
  
  
  }
  
  else {
  
  
  System.out.print(sheet.getRow(i).getCell(j).getStringCellValue());
  
  }
  
  System.out.print(" "); } System.out.println("\n");
  
  }
  
  
  workbook.close();
 

}
		
	}


