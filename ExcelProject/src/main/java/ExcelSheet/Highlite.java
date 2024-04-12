package ExcelSheet;
import org.apache.commons.compress.archivers.dump.InvalidFormatException;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined;
import org.apache.poi.ss.usermodel.*;

import jdk.internal.org.jline.utils.Log;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class Highlite {
	
	private static final Logger LOG = LogManager.getLogger(Highlite.class);
public static void main(String[] args) {
	
	 LOG.debug("This Will Be Printed On Debug");
     LOG.info("This Will Be Printed On Info");
     LOG.warn("This Will Be Printed On Warn");
     LOG.error("This Will Be Printed On Error");
     LOG.fatal("This Will Be Printed On Fatal");
     LOG.info("Appending string: {}.", "Hello, World");
	
	
	ArrayList<Integer> studentList = new ArrayList<>();
    ArrayList<Integer> gradeList = new ArrayList<>();
    ArrayList<String> header = new ArrayList<>();

    header.add("Sheet1");

    for(int i = 1; i <= 20; i++){
        studentList.add(i);
        if(i <= 20){
            gradeList.add((80+i));
        }

    }

    int bordernum = 2;
    try {
        FileOutputStream fileOut = new FileOutputStream("C:\\Users\\RK\\Desktop\\EvilTester Testcases.xlsx");
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet worksheet = workbook.createSheet("Seet1");



        // row 1 for Printing attendance sheet in center
        HSSFRow row0 = worksheet.createRow((short) 0);//1
        HSSFCell cellmid = row0.createCell((short) (gradeList.size()/2)-1);//2
        cellmid.setCellValue(header.get(0));//3
        HSSFCellStyle cellStylem = workbook.createCellStyle();//4
        cellStylem.setFillForegroundColor(HSSFColorPredefined.GOLD.getIndex());//5
        cellmid.setCellStyle(cellStylem);//6
        createBorders(workbook, cellmid, 1);
        HSSFCell cellmid2 = row0.createCell((short) (gradeList.size()/2));//2
        createBorders(workbook, cellmid2, 1);



        // row 2 with all the dates in the correct place
        HSSFRow row1 = worksheet.createRow((short) 1);//1
        HSSFCell cell1;
        for(int y = 0; y < gradeList.size(); y++){

            cell1 = row1.createCell((short) y+1);//2
            cell1.setCellValue(gradeList.get(y));//3
            createBorders(workbook, cell1, bordernum);

        }
        HSSFCellStyle cellStylei = workbook.createCellStyle();//4
        cellStylei.setFillForegroundColor(HSSFColorPredefined.GREEN.getIndex());//5



        // row 3 and on until the studentList.size() create the box.
        int counter = 0;
        for(int stu = 2; stu <= (studentList.size()+1); stu++){
            HSSFRow Row = worksheet.createRow((short) stu);//1
            for(int gr = 0; gr <= gradeList.size(); gr++){
                if(gr == 0){
                    HSSFCell cell = Row.createCell((short) 0);//2
                    cell.setCellValue(studentList.get(counter));//3
                    HSSFCellStyle cellStyle2 = workbook.createCellStyle();//4
                    cellStyle2.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
                    cellStyle2.setFillForegroundColor(HSSFColorPredefined.GOLD.getIndex());//5
                    cell.setCellStyle(cellStyle2);//6
                    createBorders(workbook, cell, 2);
                }else{
                    HSSFCell Cell = Row.createCell((short) gr);//2
                    createBorders(workbook, Cell, 3);
                }


            }
            counter++;
        }
        workbook.write(fileOut);
        fileOut.flush();
        fileOut.close();
    } catch (FileNotFoundException e) {
        e.printStackTrace();
    } catch (IOException e) {
        e.printStackTrace();
    }

}
public static void createBorders(HSSFWorkbook workbook,HSSFCell cell, int x){
    if( x == 1){
        HSSFCellStyle style = workbook.createCellStyle();
        //style.setFillBackgroundColor(HSSFColor.HSSFColorPredefined.LIGHT_BLUE.getIndex());
        //style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THICK);
        style.setBottomBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THICK);
        style.setLeftBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
        style.setBorderRight(BorderStyle.THICK);
        style.setRightBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
        style.setBorderTop(BorderStyle.THICK);
        style.setTopBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
        cell.setCellStyle(style);
    }
    else if(x == 2){
        HSSFCellStyle style = workbook.createCellStyle();
        //style.setFillBackgroundColor(HSSFColor.HSSFColorPredefined.LIGHT_BLUE.getIndex());
        //style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.MEDIUM);
        style.setBottomBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.MEDIUM);
        style.setLeftBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
        style.setBorderRight(BorderStyle.MEDIUM);
        style.setRightBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
        style.setBorderTop(BorderStyle.MEDIUM);
        style.setTopBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
        cell.setCellStyle(style);
    }else {
        HSSFCellStyle style = workbook.createCellStyle();
        //style.setFillBackgroundColor(HSSFColor.HSSFColorPredefined.AQUA.getIndex());
        //style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
        cell.setCellStyle(style);
	    }
}
		}