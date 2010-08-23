package com.miraisolutions.xlconnect;

import com.miraisolutions.xlconnect.data.DataFrame;
import com.miraisolutions.xlconnect.data.DataType;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Date;
import java.util.Vector;
import java.util.logging.LogManager;
import java.util.logging.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class App 
{
    private final static Logger logger = Logger.getLogger("com.miraisolutions.xlconnect");

    public static void main( String[] args ) throws Exception
    {
        LogManager.getLogManager().readConfiguration(App.class.getResourceAsStream("logging.properties"));

        org.apache.poi.ss.usermodel.Workbook wb = WorkbookFactory.create(new FileInputStream("C:/Users/mstuder/Documents/Test.xls"));
        wb.removeSheetAt(wb.getSheetIndex("BBB"));
        wb.write(new FileOutputStream("C:/Users/mstuder/Documents/Test2.xls"));
        if(1 == 1) return;

//        XSSFWorkbook wb = new XSSFWorkbook();
//        wb.createSheet("asdf");
//
//        DataFormat dataFormat = wb.createDataFormat();
//
//        // Header style
//        CellStyle headerStyle = wb.createCellStyle();
//        headerStyle.setDataFormat(dataFormat.getFormat("General"));
//        headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
//        headerStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
//        headerStyle.setWrapText(true);
//        // String / boolean / numeric style
//        CellStyle style = wb.createCellStyle();
//        style.setDataFormat(dataFormat.getFormat("General"));
//        style.setWrapText(true);
//        // Date style
//        CellStyle dateStyle = wb.createCellStyle();
//        dateStyle.setDataFormat(dataFormat.getFormat("mm/dd/yyyy hh:mm:ss"));
//        dateStyle.setWrapText(true);
//
//        FileOutputStream fos = new FileOutputStream("C:/temp/test.xlsx");
//        wb.write(fos);
//        fos.close();
//
//        if(1 == 1) return;
        
        File excelFile = new File("C:/temp/test.xlsx");
        if(excelFile.exists()) excelFile.delete();

        Vector<String> col1 = new Vector<String>(5);
        col1.add("A");
        col1.add("B");
        col1.add("C");
        col1.add(null);
        col1.add("E");

        Vector<Double> col2 = new Vector<Double>(5);
        for(int i = 0; i < 5; i++) col2.add(new Double(i));
        col2.setElementAt(null, 1);

        Vector<Boolean> col3 = new Vector<Boolean>(5);
        for(int i = 0; i < 5; i++) col3.add(i%2 == 0);
        col3.setElementAt(null, 2);

        Vector<Date> col4 = new Vector<Date>(5);
        for(int i = 0; i < 5; i++) col4.add(new Date(System.currentTimeMillis()));
        col4.setElementAt(null, 4);

        DataFrame df = new DataFrame();
        df.addColumn("Letter", DataType.String, col1);
        df.addColumn("Value", DataType.Numeric, col2);
        df.addColumn("Logical", DataType.Boolean, col3);
        df.addColumn("DateTime", DataType.DateTime, col4);

        Workbook workbook = Workbook.getWorkbook(excelFile, true);
        workbook.setStyleAction(StyleAction.XLCONNECT);

        // Write named region
        workbook.writeNamedRegion(df, "Test", "Test!$B$2", true);
        // Write worksheet
        workbook.writeWorksheet(df, "Test Data", 0, 0, true);
        // Add images
        workbook.addImage("C:/temp/mirai_solutions1.jpg", true, "Mirai1", "Image!$B$3:$D$5", true);
        workbook.addImage("C:/temp/mirai-solutions2.jpg", false, "Mirai2", "Image!$B$10", true);

        /** Custom style **/
        CellStyle cs = workbook.createCellStyle("MyPersonalStyle.Header");
        cs.setBorderBottom(CellStyle.BORDER_THICK);
        workbook.setStyleAction(StyleAction.STYLE_NAME_PREFIX);
        workbook.setStyleNamePrefix("MyPersonalStyle");
        workbook.writeNamedRegion(df, "Somewhere", "Somewhere!$C$5", true);

//        workbook.createCellStyle("MyStyle1");
//        workbook.createSheet("asdf");
        workbook.save();

        DataFrame res = workbook.readNamedRegion("Test", true);
        printDataFrame(res);

        res = workbook.readWorksheet("Test Data", true);
        printDataFrame(res);
    }

    public static void printDataFrame(DataFrame df) {
        for(int i = 0; i < df.columns(); i++) {
            System.out.println(df.getColumnName(i) + ":");
            System.out.println(df.getColumn(i).toString());
        }
    }
}
