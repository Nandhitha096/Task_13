
	/* Q4. Write a Java program to write data to an Excel file using Apache POI library.
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

public class QuestionFour {

     public static void main(String[] args) throws FileNotFoundException, IOException {

			//Creating object of workbook

			XSSFWorkbook wb = new XSSFWorkbook();

		    //Specifying the file path where the workbook has to be created

			String filePath = "C:\\Users\\91893\\OneDrive\\Desktop\\Student_info.xlsx";

			//Creating sheet

			XSSFSheet sh = wb.createSheet("Student_Data");
			
			//Creating ArrayList to store the list of data to the sheet
			
			ArrayList<Object[]> data = new ArrayList<Object[]>();
			
			//Using add method from ArrayList to add the required values 
			//Object[] --> holds any kind of datatypes (int, String, Boolean, etc)
			
			data.add(new Object[] {"Student_ID", "First_Name", "Last_Name", "DOB", "Percentage"} );
			data.add(new Object[] {101, "Sanket", "Singh", "09-03-1994", "93%"} );
			data.add(new Object[] {102, "Alice", "Johnson", "12-08-1997", "79%"} );
			data.add(new Object[] {103, "Smiley", "Charles", "06-06-1996", "100%"} );
			data.add(new Object[] {104, "Daisy", "Chowdry", "21-11-1996", "83%"} );
			data.add(new Object[] {105, "Sachin", "Tendulkar", "16-12-1994", "99%"} );
			data.add(new Object[] {106, "Sindhu", "Parkavi", "15-08-1996", "90%"} );
			data.add(new Object[] {107, "Divya", "Manohar", "30-10-1995", "69%"} );
			data.add(new Object[] {108, "Prakash", "Malhotra", "04-03-1994", "88%"} );
			data.add(new Object[] {109, "Dinesh", "Karthick", "11-10-1993", "61%"} );
			data.add(new Object[] {110, "Sandeep", "Singh", "29-09-1995", "45%"} );
			
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

				FileOutputStream fos = new FileOutputStream(filePath);
				wb.write(fos);
				fos.close();
				System.out.println("Given data has been written into Student_info.xlsx Excel book successfully");
			 } catch (IOException e) {
				e.printStackTrace();
			}
				
			
		}

	}

