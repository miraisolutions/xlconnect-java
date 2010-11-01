/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package com.miraisolutions.xlconnect;

/**
 *
 * @author Martin Studer, Mirai Solutions GmbH
 */
public interface CellStyle {
    void setBorderBottom(short border);
    void setBorderLeft(short border);
    void setBorderRight(short border);
    void setBorderTop(short border);
    void setBottomBorderColor(short color);
    void setLeftBorderColor(short color);
    void setRightBorderColor(short color);
    void setTopBorderColor(short color);
    void setDataFormat(String format);
    void setFillBackgroundColor(short bg);
    void setFillForegroundColor(short fp);
    void setFillPattern(short bg);
    void setWrapText(boolean wrap);
}
