/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package com.miraisolutions.xlconnect.utils;

import com.miraisolutions.xlconnect.data.DataType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;

/**
 *
 * @author Martin Studer, Mirai Solutions GmbH
 */
public abstract class CellUtils {

    public static boolean isCellValueOfType(CellValue cv, DataType type) {
        switch(cv.getCellType()) {
            case Cell.CELL_TYPE_BOOLEAN:
                return DataType.Boolean.equals(type);
            case Cell.CELL_TYPE_NUMERIC:
                return DataType.Numeric.equals(type) ||
                        (DataType.DateTime.equals(type) && DateUtil.isValidExcelDate(cv.getNumberValue()));
            case Cell.CELL_TYPE_STRING:
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
