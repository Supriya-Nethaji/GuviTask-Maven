package ReadAndWrite;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadFromFile {

	public static void main(String[] args) throws IOException {
		
		XSSFWorkbook book=new XSSFWorkbook("C:\\Users\\USER\\Desktop\\Java Practice\\GuviTaskMaven\\src\\main\\java\\ReadAndWrite\\ReadFromFile.xlsx");
		
		XSSFSheet sheet=book.getSheetAt(0);
		
		int rowCount =sheet.getLastRowNum();
		int colCount= sheet.getRow(0).getLastCellNum();
		
		//Iterating and getting the values
		for(int i=0;i<=rowCount;i++) 
		{
			XSSFRow row=sheet.getRow(i);
			
			for(int j=0;j<colCount;j++)
			{
				XSSFCell cell=row.getCell(j);
				System.out.print(cell.getStringCellValue() + " ");
			}
			System.out.println();
		}
		book.close();
	}

}
