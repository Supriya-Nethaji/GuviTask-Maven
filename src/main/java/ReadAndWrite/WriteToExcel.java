package ReadAndWrite;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteToExcel 
{
	public static void main(String args[]) throws IOException
	{
		//Creating book
		XSSFWorkbook book=new XSSFWorkbook();
		//Opening sheet
		XSSFSheet sheet = book.createSheet("Sheet 1");
		//store the data 
		Object[][] data= {
				{"Name","Age","Email"}, 
				{"John Doe",30,"John@test.com"},
				{"Jane Doe",28,"John@test.com"},
				{"Bob Smith",35,"jacky@example.com"},
				{"Swapnil",37,"swapnil@example.com"}
				};
		
		int rowCount = 0;
				
				for(Object[] row1 : data) 
				{
					XSSFRow row = sheet.createRow(rowCount++);
					
					int columnCount=0;
					for(Object col:row1) 
					{
						XSSFCell cell = row.createCell(columnCount++);
						
						if(col instanceof String) {
							cell.setCellValue((String)col);
						}else if (col instanceof Integer) {
							cell.setCellValue((Integer) col);
						}else if (col instanceof String) {
							cell.setCellValue((String) col);
						}
					}
					
				}
				
				try {
					FileOutputStream out = new FileOutputStream("C:\\Users\\USER\\Desktop\\Java Practice\\GuviTaskMaven\\src\\main\\java\\ReadAndWrite\\WriteToFile.xlsx");
					book.write(out);
				} catch (Exception e) {
					e.printStackTrace();
				}
				
				book.close();
	}
}
