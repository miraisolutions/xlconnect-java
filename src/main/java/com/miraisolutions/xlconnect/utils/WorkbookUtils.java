/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package com.miraisolutions.xlconnect.utils;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 *
 * @author mstuder
 */
public abstract class WorkbookUtils {

    public static HSSFCellStyle findCellStyleByName(HSSFWorkbook workbook, String name) {
        short count = workbook.getNumCellStyles();

        for(short i = 0; i < count; i++) {
            HSSFCellStyle cs = workbook.getCellStyleAt(i);
            if(cs.getUserStyleName().equals(name)) return cs;
        }

        return null;
    }

}
