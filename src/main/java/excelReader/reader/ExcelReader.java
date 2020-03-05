package excelReader.reader;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {

	public void read(String filePath,String sheetName) {
		try {
			FileInputStream file = new FileInputStream(new File("abc.xlsx"));
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet=null;
        if(sheetName ==null || sheetName=="" )
        	sheet= workbook.getSheetAt(0);
        else
        	sheet= workbook.getSheet(sheetName);

        Iterator<Row> rowIterator = sheet.iterator();
        while (rowIterator.hasNext()) 
        {
            Row row = rowIterator.next();
            Iterator<Cell> cellIterator = row.cellIterator();
             
            while (cellIterator.hasNext()) 
            {
                Cell cell = cellIterator.next();
                //Check the cell type and format accordingly
					switch (cell.getCellTypeEnum()) 
                {
					case NUMERIC:
                        System.out.print(cell.getNumericCellValue() + "t");
                        break;
					case STRING:
                        System.out.print(cell.getStringCellValue() + "t");
                        break;
					default:
						break;
                }
            }
            System.out.println("");
        }
        file.close();
    } 
    catch (Exception e) 
    {
        e.printStackTrace();
    }
}
}
