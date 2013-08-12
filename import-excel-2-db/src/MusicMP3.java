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

public class MusicMP3 {

	private static String excelFile = "C:/Users/dev/Desktop/test/METADATA_UPDATE.xlsx";
	
	private static String conSQL    = "jdbc:mysql://localhost:3306/music?characterEncoding=UTF-8";
	private static String userName  = "root";
	private static String password  = "123456";
	private static String tableName = "song_3";
	
	public static void main(String[] args) throws Exception {

		List<String> headers = new ArrayList<String>();
		List<List<String>> contents = new ArrayList<List<String>>();
		List<String> items = null;
		int rowCount = 0;
		int columnCount = 0;
		
		int index = 0;
		int offset = 100;
		
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
			int firstRowNum  = sheet.getFirstRowNum();
			int lastRowNum   = sheet.getLastRowNum();
			
			firstRowNum ++; // ignore header
			System.out.println("Total: " + (lastRowNum-firstRowNum) + " rows data");
			
			// get headers
			Row rowHeader = sheet.getRow(0);
			Iterator<Cell> cellIteratorHeaders = rowHeader.cellIterator();
			while (cellIteratorHeaders.hasNext()) {
				Cell cell = cellIteratorHeaders.next();
				headers.add(getCellValue(cell));			
			}
			columnCount = headers.size();

			Row row = null;
			while (index < 100) {
				for (int i = Math.max(firstRowNum, index*offset); i < Math.min(lastRowNum+1, (index+1)*offset); i++) {
					items = new ArrayList<String>();
					row = (Row)sheet.getRow(i);
					//Get iterator to all cells of current rowss
					Cell cell = null;
					for (int j = 0; j < columnCount; j++) {
						cell = row.getCell(j);
						items.add(getCellValue(cell));
					}
					contents.add(items);
				}

				int cnt = insertIntoDB(headers, contents);
				contents = new ArrayList<List<String>>();
				rowCount += cnt;
				index ++;
			}
			file.close();

			System.out.println("********************************************************");
			System.out.println("FINISH WITH: " + rowCount + " rows");
			System.out.println("********************************************************");
		} catch (Exception e) {
			System.err.println("ERROR:READ EXCEL FILE ");
			System.err.println(e.getMessage());
			throw e;
		}
	}
	
	private static int insertIntoDB(List<String> headers, List<List<String>> contents) {
		int rowCountInOnePhase = 0;
		String insertSQL = "";
		List<String> items = null;
		try {
			Class.forName("com.mysql.jdbc.Driver").newInstance();
			Connection conn = DriverManager.getConnection(conSQL, userName, password);
			Statement st = conn.createStatement();
			String insertColumns = "";
			for (int i = 0; i < headers.size(); i++) {
				insertColumns += headers.get(i);
				if ((i+1) != headers.size()) insertColumns += ",";
			}			
			
			for (int i = 1; i < contents.size(); i++) {
				items = contents.get(i);
				insertSQL = "insert into " + tableName + " ( " + insertColumns + ") values('"; 
				for (int j = 0; j < items.size(); j++) {
					insertSQL += escape(items.get(j));
					if ((j+1) != items.size()) insertSQL += "','";
				}
	
				insertSQL += "')";
				st.executeUpdate(insertSQL);
				rowCountInOnePhase ++;
			}
			System.out.println("Data is successfully inserted into the database: " + rowCountInOnePhase + " rows affected");
		} catch (Exception e) {
			System.err.println("ERROR:INSERT DATABASE ");
			System.err.println("insertSQL: " + insertSQL);
			System.err.println(e.getMessage());
		}
		return rowCountInOnePhase;
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
	
	/**
	 * Escape special characters
	 */
	private static String escape(String cell) {
		
		if (cell != null && !"".equals(cell.trim())) {
			// add more special characters here
			cell = cell.replace("\'", "\\\'");
		}
		return cell;
	}
		
}





