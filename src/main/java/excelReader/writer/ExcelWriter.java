package excelReader.writer;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriter {

	public void addSheet(String sheetName) {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet(sheetName);
		try {
			// Write the workbook in file system
			FileOutputStream out = new FileOutputStream(new File("abc.xlsx"));
			workbook.write(out);
			out.close();
			System.out.println(" Updated successfully");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void addRows(String sheetName, String excelSheetPath, String[] rows) {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet(sheetName);
		Map<String, Object[]> data = new TreeMap<String, Object[]>();
		for (int i = 0; i < rows.length; i++) {
			data.put(i + 1 + "", new Object[] { rows[i] });
		}

		// Iterate over data and write to sheet
		Set<String> keyset = data.keySet();
		int rownum = 0;
		for (String key : keyset) {
			Row row = sheet.createRow(rownum++);
			Object[] objArr = data.get(key);
			int cellnum = 0;
			for (Object obj : objArr) {
				Cell cell = row.createCell(cellnum++);
				if (obj instanceof String)
					cell.setCellValue((String) obj);
				else if (obj instanceof Integer)
					cell.setCellValue((Integer) obj);
			}
		}

		try {
			// Write the workbook in file system
			FileOutputStream out = new FileOutputStream(new File("abc.xlsx"));
			workbook.write(out);
			out.close();
			System.out.println(" Updated successfully");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
