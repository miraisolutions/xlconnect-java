/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package com.miraisolutions.xlconnect.utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellReference;

/**
 *
 * @author Martin Studer, Mirai Solutions GmbH
 */
public abstract class CellUtils {

    public static String formatAsString(Cell cell) {
        return(new CellReference(cell).formatAsString());
    }
}
