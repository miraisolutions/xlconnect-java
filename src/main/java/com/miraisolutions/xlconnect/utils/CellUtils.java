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

package com.miraisolutions.xlconnect.utils;

import com.miraisolutions.xlconnect.data.DataType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;

/**
 *
 * @author Martin Studer, Mirai Solutions GmbH
 */
public abstract class CellUtils {

    public static boolean isCellValueOfType(CellValue cv, DataType type) {
        switch(cv.getCellType()) {
            case BOOLEAN:
                return DataType.Boolean.equals(type);
            case NUMERIC:
                return DataType.Numeric.equals(type) ||
                        (DataType.DateTime.equals(type) && DateUtil.isValidExcelDate(cv.getNumberValue()));
            case STRING:
                return DataType.String.equals(type);
            default:
                return false;
        }
    }
    
    public static String formatAsString(Cell cell) {
        return(new CellReference(cell).formatAsString());
    }

    public static String getErrorMessage(FormulaError error) {
        switch(error) {
            case DIV0: return "Division by 0";
            case NA: return "Value is not available";
            case NAME: return "No such name defined";
            case NULL: return "Two areas are required to intersect but do not";
            case NUM: return "Value outside of domain";
            case REF: return "Invalid cell reference";
            case VALUE: return "Incompatible type";
            default: return "Unknown error";
        }
    }

    public static String getErrorMessage(byte errorCode) {
        return getErrorMessage(FormulaError.forInt(errorCode));
    }
}
