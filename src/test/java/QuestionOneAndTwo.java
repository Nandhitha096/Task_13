/* Q1. Create a New Excel Workbook
 * Q2. Create a New Sheet with the name "Sheet1"

 Clubbed First and Second question together in this program

 */

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class QuestionOneAndTwo {

	public static void main(String[] args) throws FileNotFoundException, IOException {

		//Creating object of workbook

		XSSFWorkbook wb = new XSSFWorkbook();

	    //Specifying the file path where the workbook has to be created

		String filePath = "C:\\Users\\91893\\OneDrive\\Desktop\\Employees.xlsx";

		//Creating sheet

		XSSFSheet sh = wb.createSheet("Sheet1");

		//Handling Exception if suppose file couldn't be created
		try {

			//Using FileOutputStream for creating file in the mentioned path
			FileOutputStream fos = new FileOutputStream(filePath);
			
			//wb.write() method writes the workbook to the file
			wb.write(fos);
			
			fos.close();
			System.out.println("The file 'Employees.xlsx' has been created with the sheet 'Sheet1' successfully");

		 } catch (IOException e) {
			e.printStackTrace();
		}

	}

}
