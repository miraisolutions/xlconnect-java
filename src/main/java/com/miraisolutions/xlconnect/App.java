/*
 *
    XLConnect
    Copyright (C) 2010 Mirai Solutions GmbH

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.
 *
 */

package com.miraisolutions.xlconnect;

import com.miraisolutions.xlconnect.data.DataFrame;
import com.miraisolutions.xlconnect.data.DataType;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.Date;
import java.util.ArrayList;
import java.util.logging.Level;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.AreaReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class App 
{
    public static void main( String[] args ) throws Exception
    {
        /* Performance measurements */
        /*
        int nrows = 10000, ncols = 100;

        File f = new File("C:/Users/mstuder/Documents/perf.xlsx");
        if(f.exists()) f.delete();

        long start = System.currentTimeMillis();

        DataFrame dfx = new DataFrame();
        for(int i = 0; i < ncols; i++) {
            ArrayList<Double> v = new ArrayList<Double>(nrows);
            for(int j = 0; j < nrows; j++)
                v.add(Math.random());

            dfx.addColumn(Integer.toString(i), DataType.Numeric, v);
        }

        long end = System.currentTimeMillis();
        System.out.println("Data creation: " + (end - start));

        Workbook wb = Workbook.getWorkbook(f, true);
        wb.createSheet("MyData");
        start = System.currentTimeMillis();
        wb.writeWorksheet(dfx, "MyData", true);
        end = System.currentTimeMillis();
        System.out.println("Write worksheet: " + (end - start));
        start = System.currentTimeMillis();
        wb.save();
        end = System.currentTimeMillis();
        System.out.println("Save: " + (end - start));
         * 
         */

        org.apache.poi.ss.usermodel.Workbook wb = WorkbookFactory.create(new FileInputStream("C:/Users/mstuder/Desktop/mtcars.xlsx"));
        Name n = wb.getName("mtcars");
        System.out.println(n.getSheetName());
        System.out.println(n.getSheetIndex());
        
        /*
        Workbook wb = Workbook.getWorkbook("C:/Users/mstuder/Desktop/mtcars.xlsx", false);
        DataFrame dfx = new DataFrame();
        ArrayList<Double> v = new ArrayList<Double>(10);
        for(int i = 0; i < 10; i++) v.add(new Double(i));
        dfx.addColumn("A", DataType.Numeric, v);
        dfx.addColumn("B", DataType.Numeric, v);
        wb.appendNamedRegion(dfx, "mtcars", false);
         * */
        
        if(1 == 1) return;

        String[] values = new String[] {"NO_FILL", "SOLID_FOREGROUND", "FINE_DOTS", "ALT_BARS", "SPARSE_DOTS",
            "THICK_HORZ_BANDS", "THICK_VERT_BANDS", "THICK_BACKWARD_DIAG", "THICK_FORWARD_DIAG", "BIG_SPOTS",
            "BRICKS", "THIN_HORZ_BANDS", "THIN_VERT_BANDS", "THIN_BACKWARD_DIAG", "THIN_FORWARD_DIAG",
            "SQUARES", "DIAMONDS"};

        Class c = Class.forName("org.apache.poi.ss.usermodel.CellStyle");
        for(String val : values) {
            System.out.println("XLC$\"FILL." + val + "\" <- " + c.getField(val).get(null));
        }

        org.apache.poi.ss.usermodel.CellStyle css;

        if(1 == 1) return;

        File excelFile = new File("C:/temp/test.xlsx");
        if(excelFile.exists()) excelFile.delete();

        ArrayList<String> col1 = new ArrayList<String>(5);
        col1.add("A");
        col1.add("B");
        col1.add("C");
        col1.add(null);
        col1.add("E");

        ArrayList<Double> col2 = new ArrayList<Double>(5);
        for(int i = 0; i < 5; i++) col2.add(new Double(i));
        col2.set(1, null);

        ArrayList<Boolean> col3 = new ArrayList<Boolean>(5);
        for(int i = 0; i < 5; i++) col3.add(i%2 == 0);
        col3.set(2, null);

        ArrayList<Date> col4 = new ArrayList<Date>(5);
        for(int i = 0; i < 5; i++) col4.add(new Date(System.currentTimeMillis()));
        col4.set(4, null);

        DataFrame df = new DataFrame();
        df.addColumn("Letter", DataType.String, col1);
        df.addColumn("Value", DataType.Numeric, col2);
        df.addColumn("Logical", DataType.Boolean, col3);
        df.addColumn("DateTime", DataType.DateTime, col4);

        Workbook workbook = Workbook.getWorkbook(excelFile, true);
        workbook.setStyleAction(StyleAction.XLCONNECT);

        // Write named region
        workbook.createName("Test", "Test!$B$2", true);
        workbook.writeNamedRegion(df, "Test", true);

        //--
        // Write worksheet
        workbook.createSheet("Test Data");
        workbook.writeWorksheet(df, "Test Data", 0, 0, true);
        // Add images
        workbook.createName("Mirai1", "Image!$B$3:$D$5", true);
        workbook.createName("Mirai2", "Image!$B$10", true);
        workbook.addImage("C:/temp/mirai_solutions1.jpg", "Mirai1", true);
        workbook.addImage("C:/temp/mirai-solutions2.jpg", "Mirai2", false);

        
        CellStyle cs = workbook.createCellStyle("MyPersonalStyle.Header");
        cs.setBorderBottom(org.apache.poi.ss.usermodel.CellStyle.BORDER_THICK);
        // cs.setFillPattern(org.apache.poi.ss.usermodel.CellStyle.SOLID_FOREGROUND);
        // cs.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        workbook.setStyleAction(StyleAction.STYLE_NAME_PREFIX);
        workbook.setStyleNamePrefix("MyPersonalStyle");
        workbook.createName("Somewhere", "Somewhere!$C$5", true);
        workbook.writeNamedRegion(df, "Somewhere", true);     

        CellStyle funky = workbook.createCellStyle();
        funky.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
        funky.setFillPattern(org.apache.poi.ss.usermodel.CellStyle.SOLID_FOREGROUND);
        workbook.setCellStyle("Somewhere!$D$6:$E$9", funky);
        
        // ---

        workbook.save();

        /**
        DataFrame res = workbook.readNamedRegion("Test", true);
        printDataFrame(res);

        res = workbook.readWorksheet("Test Data", true);
        printDataFrame(res);
         **/
    }

    public static void printDataFrame(DataFrame df) {
        for(int i = 0; i < df.columns(); i++) {
            System.out.println(df.getColumnName(i) + ":");
            System.out.println(df.getColumn(i).toString());
        }
    }
}
