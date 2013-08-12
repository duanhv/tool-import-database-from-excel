import java.io.File;
import java.io.FileInputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GenTableScript {

	private static String excelFile = "C:/Users/dev/Desktop/test/ToadExport_test.xlsx";	
	private static String tableName = "song_3";
	
	public static void main(String[] args) throws Exception {

		List<String> headers = new ArrayList<String>();

		try {
			//..
			FileInputStream file = new FileInputStream(new File(excelFile));
			System.out.println("STEP1");
			//Get the workbook instance for XLS file
			XSSFWorkbook workbook = new XSSFWorkbook (file);
			System.out.println("STEP2");
			//Get first sheet from the workbook
			XSSFSheet sheet = workbook.getSheetAt(0);
			System.out.println("STEP3");

			// get headers
			Row rowHeader = sheet.getRow(0);
			Iterator<Cell> cellIteratorHeaders = rowHeader.cellIterator();
			while (cellIteratorHeaders.hasNext()) {
				Cell cell = cellIteratorHeaders.next();
				headers.add(getCellValue(cell));			
			}

			String script = "CREATE TABLE `" + tableName + "` ("; 
			script += "`ID` int(11) NOT NULL AUTO_INCREMENT," + "\n";
			for (int i = 1; i < headers.size(); i++) {

				script += "`" + headers.get(i) + "` varchar(200) DEFAULT NULL," + "\n";

			}
			script += " PRIMARY KEY (`ID`) ";
			script += ") ENGINE=InnoDB DEFAULT CHARSET=utf8;";
			
			file.close();

			System.out.println("********************************************************");
			System.err.println(script);
			System.out.println("********************************************************");
		} catch (Exception e) {
			System.err.println("ERROR:READ EXCEL FILE ");
			System.err.println(e.getMessage());
			throw e;
		}
	}

	/**
	 * Check cell type to get data
	 */
	private static String getCellValue(Cell cell) {
	
		String res = "";
		//System.out.println("cell.getCellType() = " + cell.getCellType());
		if (cell != null) { 
			switch (cell.getCellType()) {
			case Cell.CELL_TYPE_STRING:
				res = cell.getStringCellValue();
				break;
			case Cell.CELL_TYPE_NUMERIC:
				res = String.valueOf(cell.getNumericCellValue());
				break;
			case Cell.CELL_TYPE_BOOLEAN:
				res = String.valueOf(cell.getBooleanCellValue());
				break;
			}
		}
		return res;
	}

}





