/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package com.miraisolutions.xlconnect.utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.util.CellReference;

/**
 *
 * @author Martin Studer, Mirai Solutions GmbH
 */
public abstract class CellUtils {

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
