package CBP.SeleniumExcel;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class dataDriven1 {
	public ArrayList getdata(String testCaseName) throws IOException {
		// Identify the testcase column by scanning the entire row
		// once column is identified and scan the entire test case for purchase test
		// case
		// after you grab the purchase test case row =pull all the data of that row and
		// feed it into test

		ArrayList<String> a = new ArrayList<String>();
		// Construct the file path
		String filePath = System.getProperty("user.dir") + "//ExcelData//Book1.xlsx";

		// Create FileInputStream object
		FileInputStream fis = new FileInputStream(filePath);

		// Create XSSFWorkbook object
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		// First get the number of sheets in the excel
		int sheets = workbook.getNumberOfSheets();
		for (int i = 0; i < sheets; i++) {
			if (workbook.getSheetName(i).equalsIgnoreCase("Testdata")) {
				XSSFSheet sheet = workbook.getSheetAt(i);
				// sheet is a collection of rows
				Iterator<Row> rows = sheet.iterator();
				Row firstRow = rows.next();
				// row is a collection of cells
				Iterator<Cell> ce = firstRow.cellIterator();
				int k = 0;
				int column = 0;
				while (ce.hasNext()) {
					Cell value = ce.next();
					if (value.getStringCellValue().equalsIgnoreCase(testCaseName)) {
						column = k;
					}
					k++;
				}
				System.out.println(column);
				// Once column is identified scan the entire column to identify the purchase
				// testcase row
				while (rows.hasNext()) {
					Row r = rows.next();
					if (r.getCell(column).getStringCellValue().equalsIgnoreCase(testCaseName)) {
						Iterator<Cell> cv = r.cellIterator();
						while (cv.hasNext()) {
							Cell c = cv.next();
							if (c.getCellType() == CellType.STRING) {
								a.add(c.getStringCellValue());
							} else {
								a.add(NumberToTextConverter.toText(c.getNumericCellValue()));

							}

						}
					}
				}
			}
		}

		return a;
	}

//Identify the testcase column by scanning the entire row
	// once column is identified and scan the entire test case for purchase test
	// case
	// after you grab the purchase test case row =pull all the data of that row and
	// feed it into test
	public static void main(String[] args) throws IOException {
		dataDriven1 d = new dataDriven1();
		ArrayList<String> data = d.getdata("Add Profile");
		System.out.println(data.get(0));
		System.out.println(data.get(1));
		System.out.println(data.get(2));
		System.out.println(data.get(3));
	}
}
