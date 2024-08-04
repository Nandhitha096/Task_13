import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class QuestionFive {


	public static void main(String[] args) throws IOException {
		
		
		//Specifying the location of Excel file
		File src = new File("C:\\Users\\91893\\OneDrive\\Desktop\\Student_info.xlsx");
		
		//Loading the file 
		FileInputStream fis = new FileInputStream(src);
	    
		//Loading the workbook 
		XSSFWorkbook wb= new XSSFWorkbook(fis);
		
		//Loading the sheet 
		XSSFSheet sh = wb.getSheet("Student_Data");
			
		//Fetching the number of rows and columns in the sheet
		int rows = sh.getLastRowNum()+1;
		int columns = sh.getRow(0).getLastCellNum();
		
		//Printing the total  number of rows and columns
		System.out.println("Total number of Rows : "+rows);
		System.out.println("Total number of columns : "+columns);
		
		for(int i=0;i<rows;i++)
		{
			for(int j=0;j<columns;j++) 
			{
				System.out.println(sh.getRow(i).getCell(j).getStringCellValue() + "\t");
				
			}
			
			System.out.println("\n");
		}
		
}
		
}