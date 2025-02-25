/*
 *
    XLConnect
    Copyright (C) 2010-2024 Mirai Solutions GmbH

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

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;

public final class HCellStyle implements CellStyle {

    private final HSSFWorkbook workbook;
    private final HSSFCellStyle cellStyle;

    public HCellStyle(HSSFWorkbook workbook, HSSFCellStyle cellStyle) {
        this.workbook = workbook;
        this.cellStyle = cellStyle;
    }

    public void setBorderBottom(BorderStyle border) {
        cellStyle.setBorderBottom(border);
    }

    public void setBorderLeft(BorderStyle border) {
        cellStyle.setBorderLeft(border);
    }

    public void setBorderRight(BorderStyle border) {
        cellStyle.setBorderRight(border);
    }

    public void setBorderTop(BorderStyle border) {
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

    public void setFillPattern(FillPatternType bg) {
        cellStyle.setFillPattern(bg);
    }

    public void setWrapText(boolean wrap) {
        cellStyle.setWrapText(wrap);
    }

    public void setFontBold(boolean bold) {
        throw new UnsupportedOperationException("Font operations are not supported in HCellStyle.");
    }
    public void setFontName(String name) {
        throw new UnsupportedOperationException("Font operations are not supported for SSCellStyle.");
    }

    public void setFontItalic(boolean italic) {
        throw new UnsupportedOperationException("Font operations are not supported for SSCellStyle.");
    }

    public void setFontSize(int size) {
        throw new UnsupportedOperationException("Font operations are not supported for SSCellStyle.");
    }
    public void setFontColor(short color, byte[] rgb) {
        throw new UnsupportedOperationException("Font color is not supported for HCellStyle.");
    }

    public static HCellStyle create(HSSFWorkbook workbook, String name) {
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        if (name != null) cellStyle.setUserStyleName(name);
        return new HCellStyle(workbook, cellStyle);
    }

    public static HCellStyle get(HSSFWorkbook workbook, String name) {
        for (short i = 0; i < workbook.getNumCellStyles(); i++) {
            HSSFCellStyle cs = workbook.getCellStyleAt(i);
            String userStyleName = cs.getUserStyleName();
            if (userStyleName != null && cs.getUserStyleName().equals(name))
                return new HCellStyle(workbook, cs);
        }

        return null;
    }

    public static void set(HSSFCell c, HCellStyle cs) {
        c.setCellStyle(cs.cellStyle);
    }
}
