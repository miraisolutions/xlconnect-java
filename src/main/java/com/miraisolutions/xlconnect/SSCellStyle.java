/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package com.miraisolutions.xlconnect;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormat;

/**
 *
 * @author Martin Studer, Mirai Solutions GmbH
 */
public class SSCellStyle implements CellStyle {

    org.apache.poi.ss.usermodel.Workbook workbook;
    org.apache.poi.ss.usermodel.CellStyle cellStyle;

    public SSCellStyle(org.apache.poi.ss.usermodel.Workbook workbook, org.apache.poi.ss.usermodel.CellStyle cellStyle) {
        this.workbook = workbook;
        this.cellStyle = cellStyle;
    }

    public void setBorderBottom(short border) {
        cellStyle.setBorderBottom(border);
    }

    public void setDataFormat(String format) {
        DataFormat dataFormat = workbook.createDataFormat();
        cellStyle.setDataFormat(dataFormat.getFormat(format));
    }

    public void setFillForegroundColor(short fp) {
        cellStyle.setFillForegroundColor(fp);
    }

    public void setFillPattern(short bg) {
        cellStyle.setFillPattern(bg);
    }

    public void setWrapText(boolean wrap) {
        cellStyle.setWrapText(wrap);
    }

    public static void set(Cell c, SSCellStyle cs) {
        c.setCellStyle(cs.cellStyle);
    }
}
