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

import com.miraisolutions.xlconnect.data.DataType;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;

import java.util.EnumMap;

/**
 * Marker cell style used to specify that a cell style
 * should be determined dynamically with the data format
 * being re-specified according to the data type
 */
public class DataFormatOnlyCellStyle extends Common implements CellStyle {

    private static final EnumMap<DataType, DataFormatOnlyCellStyle> instances = new EnumMap<>(DataType.class);
    private final DataType dataType;
    
    private DataFormatOnlyCellStyle(DataType type) {
        this.dataType = type;
    }

    public DataType getDataType() {
        return dataType;
    }

    public void setBorderBottom(BorderStyle border) {
        throw new UnsupportedOperationException();
    }

    public void setBorderLeft(BorderStyle border) {
        throw new UnsupportedOperationException();
    }

    public void setBorderRight(BorderStyle border) {
        throw new UnsupportedOperationException();
    }

    public void setBorderTop(BorderStyle border) {
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

    public void setFillPattern(FillPatternType bg) {
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
    
    public String getDataFormat() {
        throw new UnsupportedOperationException();
    }

    public static DataFormatOnlyCellStyle get(DataType type) {
        DataFormatOnlyCellStyle cs;
        if(instances.containsKey(type))
            cs = instances.get(type);
        else {
            cs = new DataFormatOnlyCellStyle(type);
            instances.put(type, cs);
        }
        return cs;
    }
}
