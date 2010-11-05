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

    public void setBorderLeft(short border) {
        cellStyle.setBorderLeft(border);
    }

    public void setBorderRight(short border) {
        cellStyle.setBorderRight(border);
    }

    public void setBorderTop(short border) {
        cellStyle.setBorderTop(border);
    }

    public void setBottomBorderColor(short color) {
        cellStyle.setBottomBorderColor(color);
    }

    public void setLeftBorderColor(short color) {
        cellStyle.setLeftBorderColor(color);
    }

    public void setRightBorderColor(short color) {
        cellStyle.setRightBorderColor(color);
    }

    public void setTopBorderColor(short color) {
        cellStyle.setTopBorderColor(color);
    }

    public void setDataFormat(String format) {
        DataFormat dataFormat = workbook.createDataFormat();
        cellStyle.setDataFormat(dataFormat.getFormat(format));
    }

    public void setFillBackgroundColor(short bg) {
        cellStyle.setFillBackgroundColor(bg);
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
