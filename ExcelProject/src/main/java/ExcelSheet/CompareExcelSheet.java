package ExcelSheet;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class CompareExcelSheet {

	    public static void main(String[] args) {
	        try {
	            // Path to your dynamic and static Excel files
	            String saucedemo = "C:\\Users\\RK\\Desktop\\Saucedemo.com.xlsx";
	            String eviltester = "C:\\Users\\RK\\Desktop\\EvilTester Testcases.xlsx";

	            // Load both Excel files
	            Workbook saucedemoBranch = WorkbookFactory.create(new FileInputStream(new File(saucedemo)));
	            Workbook eviltesterMaster= WorkbookFactory.create(new FileInputStream(new File(eviltester)));

	            // Compare each sheet
	            for (int i = 0; i < saucedemoBranch.getNumberOfSheets(); i++) {
	                Sheet dynamicSheetS = saucedemoBranch.getSheetAt(i);
	                Sheet staticSheetA = eviltesterMaster.getSheetAt(i);

	                // Compare each row in the sheet
	                for (int j = 0; j <= dynamicSheetS.getLastRowNum(); j++) {
	                    Row saucedemoRow = dynamicSheetS.getRow(j);
	                    Row eviltesterRow = staticSheetA.getRow(j);

	                    // Compare each cell in the row
	                    for (int k = 0; k < saucedemoRow.getLastCellNum(); k++) {
	                        Cell dynamicCell = eviltesterRow.getCell(i);
	                        Cell staticCell = saucedemoRow.getCell(k);

	                        // Compare cell values
	                        if (dynamicCell != null && staticCell != null) {
	                            if (!dynamicCell.toString().equals(staticCell.toString())) {
	                                System.out.println("Mismatch found at Sheet: " + dynamicSheetS.getSheetName() +
	                                        ", Row: " + (k + 1) + ", Column: " + (i + 1));
	                            }
	                        } else if (dynamicCell == null && staticCell != null) {
	                            System.out.println("Mismatch found at Sheet: " + dynamicSheetS.getSheetName() +
	                                    ", Row: " + (k + 1) + ", Column: " + (i + 1));
	                        } else if (dynamicCell != null && staticCell == null) {
	                            System.out.println("Mismatch found at Sheet: " + dynamicSheetS.getSheetName() +
	                                    ", Row: " + (k + 1) + ", Column: " + (i + 1));
	                        }
	                    }
	                }
	            }

	            // Close the workbooks
	            saucedemoBranch.close();
	            eviltesterMaster.close();
	            }
	        
	   catch (Exception e) {
		// TODO: handle exception
	} 

}}
