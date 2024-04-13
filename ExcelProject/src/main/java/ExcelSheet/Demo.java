package ExcelSheet;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Demo {

	public static void main(String[] args) throws IOException {
		
		
		
	            
	       
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet= workbook.createSheet("Sheet1");
		XSSFRow row=sheet.createRow(1);
		
	
		
		XSSFCellStyle style=workbook.createCellStyle(); style.setFillBackgroundColor (IndexedColors.BRIGHT_GREEN.getIndex()); style.setFillPattern (FillPatternType.BIG_SPOTS);

		XSSFCell cell= row.createCell(1); 
		cell.setCellValue("welcome"); 
		cell.setCellStyle(style);

		// Setting Foreground color

		style=workbook.createCellStyle();

		style.setFillForegroundColor(IndexedColors. YELLOW.getIndex()); 
		
		style.setFillPattern (FillPatternType.SOLID_FOREGROUND);

		cell= row.createCell(1);

		cell.setCellValue("welcome");

		cell.setCellStyle(style);
		FileOutputStream file = new FileOutputStream("C:\\\\Users\\\\RK\\\\Desktop\\\\Demo.xlsx");
		 
		workbook.write(file);
		workbook.close();

		file.close();

		System.out.println("Done!!!");
		
		
//		
		        
        try {
    // Load the Excel files
    FileInputStream fis1 = new FileInputStream("C:\\Users\\RK\\Desktop\\EvilTester Testcases.xlsx");
    FileInputStream fis2 = new FileInputStream("C:\\Users\\RK\\Desktop\\Saucedemo.com.xlsx");
    Workbook workbook1 = WorkbookFactory.create(fis1);
    Workbook workbook2 = WorkbookFactory.create(fis2);
    
    // Create a new Excel workbook to highlight the differences
    Workbook diffWorkbook = new XSSFWorkbook();
    org.apache.poi.ss.usermodel.Sheet diffSheet=diffWorkbook.createSheet("Differences");
    
    // Iterate through each sheet and cell
    for (int sheetIndex = 0; sheetIndex < workbook1.getNumberOfSheets(); sheetIndex++) {
        org.apache.poi.ss.usermodel.Sheet sheet1 = workbook1.getSheetAt(sheetIndex);
        org.apache.poi.ss.usermodel.Sheet sheet2 = workbook2.getSheetAt(sheetIndex);
        org.apache.poi.ss.usermodel.Sheet diffSheet1 = diffWorkbook.createSheet(sheet1.getSheetName());
        
        for (int rowIndex = 0; rowIndex <= sheet1.getLastRowNum(); rowIndex++) {
            Row row1 = sheet1.getRow(rowIndex);
            Row row2 = sheet2.getRow(rowIndex);
            Row diffRow = diffSheet1.createRow(rowIndex);
            
            for (int cellIndex = 0; cellIndex < row1.getLastCellNum(); cellIndex++) {
                Cell cell1 = row1.getCell(cellIndex);
                Cell cell2 = row2.getCell(cellIndex);
                Cell diffCell = diffRow.createCell(cellIndex);
                
                if (cell1 == null || cell2 == null) {
                    continue;
                }
                
                if (!cell1.equals(cell2)) {
                    // Apply formatting to highlight the difference
                    CellStyle style1 = diffWorkbook.createCellStyle();
                    style1.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
                    style1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    diffCell.setCellStyle(style1);
                }
                
                // Copy cell value to the diff sheet
                diffCell.setCellValue(cell1.toString());
            }
        }
    }
    
    // Write the differences to a new Excel file
    FileOutputStream fos = new FileOutputStream("differences.xlsx");
    diffWorkbook.write(fos);
    fos.close();
    
    // Close all workbooks
    workbook1.close();
    workbook2.close();
    diffWorkbook.close();
    
    System.out.println("Differences highlighted successfully!");
} catch (Exception e) {
    System.err.println("An error occurred: " + e.getMessage());
}
       
    
//
		
	}

}
