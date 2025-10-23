import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelAccess {
	public static void main(String[] args) throws IOException {
	FileInputStream fis = new FileInputStream("/Users/prajwalbodhale/eclipse-workspace/ExcelDriven/EmployeeData.xlsx");
	XSSFWorkbook workbook = new XSSFWorkbook(fis);
	DataFormatter formatter =new DataFormatter();
	XSSFSheet sheet = workbook.getSheetAt(0);

	
	Iterator<Row> rowIterator = sheet.iterator(); // sheet is collection of rows
	
	//***To get Specific row data with e.g firstcell = employeeid *******
	String employee_id = "E4015";
	String headings = "EmployeeID";
	while(rowIterator.hasNext()) {
		Row row = rowIterator.next();
		Iterator<Cell> cellIterator = row.cellIterator();
		if(cellIterator.hasNext()) {
			Cell value = cellIterator.next();
			String cellValue = formatter.formatCellValue(value);
			if(cellValue.equals(headings)) {
				System.out.println(cellValue + "|");
				while(cellIterator.hasNext()) {
					Cell nextCell = cellIterator.next();
					System.out.println(formatter.formatCellValue(nextCell)+"|");
				}
				System.out.println();
				break;
				}
			
			}
		}
	while(rowIterator.hasNext()) {
		Row row = rowIterator.next();
		Iterator<Cell> cellIterator = row.cellIterator();
		if(cellIterator.hasNext()) {
			Cell value = cellIterator.next();
			String cellValue = formatter.formatCellValue(value);
			if(cellValue.equals(employee_id)) {
				System.out.println(cellValue + "|");
				while(cellIterator.hasNext()) {
					Cell nextCell = cellIterator.next();
					System.out.println(formatter.formatCellValue(nextCell)+"|");
				}
				break;
				}
			
			}
		}
	
	}
}
