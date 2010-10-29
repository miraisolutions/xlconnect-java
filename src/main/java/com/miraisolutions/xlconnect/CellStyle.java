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
    void setDataFormat(String format);
    void setFillForegroundColor(short fp);
    void setFillPattern(short bg);
    void setWrapText(boolean wrap);
}
