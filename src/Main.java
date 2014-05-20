import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class Main {
	private static final int MY_MINIMUM_COLUMN_COUNT = 0;

	public static void main(String[] args) throws InvalidFormatException, IOException {
		Main main = new Main();
		main.read();
	}
	
	void read()
			throws InvalidFormatException, IOException {
		InputStream inp = new FileInputStream("input.xlsx");
		Workbook wb = WorkbookFactory.create(inp);
		int sheetCount = wb.getNumberOfSheets();
		for (int i = 0; i < sheetCount; i++) {

			Sheet sheet = wb.getSheetAt(i);
			Workbook wbOutput = new XSSFWorkbook();
			Sheet sheetOutput = wbOutput.createSheet();
			int rowStart = sheet.getFirstRowNum();
			int rowEnd = sheet.getLastRowNum() + 1;
			// 读取矩阵
			for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
				Row r = sheet.getRow(rowNum);
				Row row = sheetOutput.createRow(rowNum);
				int lastColumn = Math.max(r.getLastCellNum(),
						MY_MINIMUM_COLUMN_COUNT);
				for (int cn = 0; cn < lastColumn; cn++) {
					Cell c = r.getCell(cn, Row.RETURN_BLANK_AS_NULL);
					if (c == null) {
						row.createCell(cn).setCellValue("");
					} else {
						row.createCell(cn).setCellValue(c.toString());
					}
				}
			}
			FileOutputStream fileOut = new FileOutputStream(sheet.getSheetName() + ".xlsx");
			wbOutput.write(fileOut);
			fileOut.close();
		}
	}
}
