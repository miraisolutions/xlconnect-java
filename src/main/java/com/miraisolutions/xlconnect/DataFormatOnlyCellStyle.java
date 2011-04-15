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

/**
 * Marker cell style used to specify that a cell style
 * should be determined dynamically with the data format
 * being re-specified according to the data type
 * 
 * @author Martin Studer, Mirai Solutions GmbH
 */
public class DataFormatOnlyCellStyle implements CellStyle {

    private static DataFormatOnlyCellStyle instance = null;

    public void setBorderBottom(short border) {
        throw new UnsupportedOperationException();
    }

    public void setBorderLeft(short border) {
        throw new UnsupportedOperationException();
    }

    public void setBorderRight(short border) {
        throw new UnsupportedOperationException();
    }

    public void setBorderTop(short border) {
        throw new UnsupportedOperationException();
    }

    public void setBottomBorderColor(short color) {
        throw new UnsupportedOperationException();
    }

    public void setDataFormat(String format) {
        throw new UnsupportedOperationException();
    }

    public void setFillBackgroundColor(short bg) {
        throw new UnsupportedOperationException();
    }

    public void setFillForegroundColor(short fp) {
        throw new UnsupportedOperationException();
    }

    public void setFillPattern(short bg) {
        throw new UnsupportedOperationException();
    }

    public void setLeftBorderColor(short color) {
        throw new UnsupportedOperationException();
    }

    public void setRightBorderColor(short color) {
        throw new UnsupportedOperationException();
    }

    public void setTopBorderColor(short color) {
        throw new UnsupportedOperationException();
    }

    public void setWrapText(boolean wrap) {
        throw new UnsupportedOperationException();
    }

    public static DataFormatOnlyCellStyle get() {
        if(instance == null)
            instance = new DataFormatOnlyCellStyle();
        return instance;
    }
}
