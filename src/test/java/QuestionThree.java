
	/* Q3. Write the data into the sheet
	 */

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class QuestionThree {

     public static void main(String[] args) throws FileNotFoundException, IOException {

			//Creating object of workbook

			XSSFWorkbook wb = new XSSFWorkbook();

		    //Specifying the file path where the workbook has to be created

			String filePath = "C:\\Users\\91893\\OneDrive\\Desktop\\Employees.xlsx";

			//Creating sheet

			XSSFSheet sh = wb.createSheet("Sheet1");
			
			//Creating ArrayList to store the list of data to the sheet
			
			ArrayList<Object[]> data = new ArrayList<Object[]>();
			
			//Using add method from ArrayList to add the required values 
			//Object[] --> holds any kind of datatypes (int, String, Boolean, etc)
			
			data.add(new Object[] {"Name", "Age", "Email"} );
			data.add(new Object[] {"John Doe", 30, "john@test.com"} );
			data.add(new Object[] {"Jane Doe", 28, "john@test.com"} );
			data.add(new Object[] {"Bob Smith", 35, "jacky@example.com"} );
			data.add(new Object[] {"Swapnil", 37, "swapnil@example.com"} );
			
			int row=0;
			
			//for each loop to write the values to the sheet in form of table
			for(Object[] i:data) {
				
				XSSFRow r = sh.createRow(row++);
				int column=0;
				
				for(Object value:i)
				{
					
					XSSFCell c = r.createCell(column++);
					
					//Checking the type of datatype and typecasting
					if(value instanceof String)
						c.setCellValue((String) value);
					else if(value instanceof Integer)
						c.setCellValue((Integer) value);
					else if(value instanceof Boolean)
						c.setCellValue((Boolean) value);
				}
				
			}
			
			//Handling Exception if suppose file couldn't be created
			try {

				//Writing to file
				FileOutputStream fos = new FileOutputStream(filePath);
				wb.write(fos);
				fos.close();
				System.out.println("Given data has been written into Employee.xlsx worksheet successfully");
			 } catch (IOException e) {
				e.printStackTrace();
			}
				
			
		}

	}

