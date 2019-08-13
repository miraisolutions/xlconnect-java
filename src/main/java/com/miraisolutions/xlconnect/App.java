/*
 *
    XLConnect
    Copyright (C) 2010-2018 Mirai Solutions GmbH

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

import com.miraisolutions.xlconnect.data.Column;
import com.miraisolutions.xlconnect.data.DataFrame;
import com.miraisolutions.xlconnect.data.DataType;
import com.miraisolutions.xlconnect.data.ReadStrategy;
import com.miraisolutions.xlconnect.integration.r.RDataFrameWrapper;
import com.miraisolutions.xlconnect.integration.r.RWorkbookWrapper;
import com.miraisolutions.xlconnect.utils.RPOSIXDateTimeFormatter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.lang.reflect.Array;
import java.lang.reflect.Field;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.ArrayList;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.SpreadsheetVersion;
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
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class App 
{
    public static void main( String[] args ) throws Exception
    {
        String file = "/home/mstuder/Downloads/SEA 05 Staffing Call 2 Test.xlsx";
        org.apache.poi.ss.usermodel.Workbook wb = new XSSFWorkbook(file);
        Name name = wb.getName("Data");
        AreaReference ref = AreaReference.generateContiguous(wb.getSpreadsheetVersion(), name.getRefersToFormula())[0];

        // AreaReference ref = new AreaReference(name.getRefersToFormula(), SpreadsheetVersion.EXCEL2007);


      /*  File f = new File(file);
        if(f.exists()) f.delete();
        Workbook wb = Workbook.getWorkbook(f, true);
        wb.setStyleAction(StyleAction.DATATYPE);
        CellStyle cs = wb.createCellStyle();
        cs.setDataFormat("d/m/yy");
        wb.setCellStyleForDataType(DataType.DateTime, cs);
        DataFrame df = new DataFrame();
        boolean[] missing = new boolean[] {false, false, false, false, false};
        Date date = new Date();
        df.addColumn("A", new Column(new double[] {1.0, 2.0, 3.0, 4.0, 5.0}, missing, DataType.Numeric));
        df.addColumn("B", new Column(new Date[] {date, date, date, date, date}, missing, DataType.DateTime));
        wb.createSheet("data");
        wb.writeWorksheet(df, "data", true);
        wb.save();
        printDataFrame(df);*/
        
    }

    public static void printDataFrame(DataFrame df) {
        for(int i = 0; i < df.columns(); i++) {
            System.out.println(df.getColumnName(i) + ":");
            Object data = df.getColumn(i).getData();
            boolean[] missing = df.getColumn(i).getMissing();
            int len = missing.length;
            for(int j = 0; j < len; j++) {
                if(missing[j])
                    System.out.print("[NA]");
                else
                    System.out.print(Array.get(data, j));
                System.out.print(" ");
            }
            System.out.println();
        }
    }
}
