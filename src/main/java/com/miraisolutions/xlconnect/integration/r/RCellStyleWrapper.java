/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package com.miraisolutions.xlconnect.integration.r;

import com.miraisolutions.xlconnect.CellStyle;

/**
 *
 * @author Martin Studer, Mirai Solutions GmbH
 */
public class RCellStyleWrapper {

    final CellStyle cellStyle;

    public RCellStyleWrapper(CellStyle cellStyle) {
        this.cellStyle = cellStyle;
    }

    public void setBorderBottom(int border) {
        cellStyle.setBorderBottom((short) border);
    }

    void setDataFormat(String format) {
        cellStyle.setDataFormat(format);
    }

    void setFillForegroundColor(int fp) {
        cellStyle.setFillForegroundColor((short) fp);
    }

    void setFillPattern(int bg) {
        cellStyle.setFillPattern((short) bg);
    }

    void setWrapText(boolean wrap) {
        cellStyle.setWrapText(wrap);
    }
}
