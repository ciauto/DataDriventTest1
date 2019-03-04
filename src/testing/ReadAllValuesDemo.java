package testing;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadAllValuesDemo {

	public static void main(String[] args) throws IOException {
		// specify the excel file containing test data
		File src = new File("C:\\Users\\Naresh\\oxygen-workspace\\DataDrivenTest1\\testData.xlsx");
		// load the excel file
		FileInputStream fis = new FileInputStream(src);
		// load the workbook from the above excel file
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		// load the sheet from above excel workbook
		XSSFSheet sheet = wb.getSheetAt(0);
		int rowCount=sheet.getLastRowNum();
		for(int i=1; i<=rowCount; i++) {
			//Reading first column(getCell refer to Column)
			String data=sheet.getRow(i).getCell(0).getStringCellValue();
			System.out.println("Excel data is:  "+data);
		}
		// close the workbook
		wb.close();
	}
}
