/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package com.miraisolutions.xlconnect;

/**
 *
 * @author Martin Studer, Mirai Solutions GmbH
 */
public class SSCellStyle implements CellStyle {

    org.apache.poi.ss.usermodel.CellStyle cellStyle;

    public SSCellStyle(org.apache.poi.ss.usermodel.CellStyle cellStyle) {
        this.cellStyle = cellStyle;
    }

    public org.apache.poi.ss.usermodel.CellStyle getPOICellStyle() {
        return cellStyle;
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

}
