/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package com.miraisolutions.xlconnect;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 *
 * @author Martin Studer, Mirai Solutions GmbH
 */
public class HCellStyle implements CellStyle {

    private final HSSFCellStyle cellStyle;

    public HCellStyle(HSSFCellStyle cellStyle) {
        this.cellStyle = cellStyle;
    }
    
    public void setBorderBottom(short border) {
        cellStyle.setBorderBottom(border);
    }

    public void setDataFormat(short format) {
        cellStyle.setDataFormat(format);
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

    public static HCellStyle create(HSSFWorkbook workbook, String name) {
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        if(name != null) cellStyle.setUserStyleName(name);
        return new HCellStyle(cellStyle);
    }

    public static HCellStyle get(HSSFWorkbook workbook, String name) {
        HSSFWorkbook wb = (HSSFWorkbook) workbook;
        for(short i = 0; i < workbook.getNumCellStyles(); i++) {
            HSSFCellStyle cs = wb.getCellStyleAt(i);
            String userStyleName = cs.getUserStyleName();
            if(userStyleName != null && cs.getUserStyleName().equals(name))
                return new HCellStyle(cs);
        }

        return null;
    }

    public static void set(HSSFCell c, HCellStyle cs) {
        c.setCellStyle(cs.cellStyle);
    }
}
