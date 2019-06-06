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

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellAlignment;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorder;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorderPr;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCellAlignment;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCellStyle;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCellStyles;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPatternFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTXf;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STBorderStyle;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STPatternType;

/**
 * This class uses parts from XSSFCellStyle.java at
 * http://svn.apache.org/repos/asf/poi/trunk/src/ooxml/java/org/apache/poi/xssf/usermodel/XSSFCellStyle.java
 * by The Apache Software Foundation
 */
public class XCellStyle extends Common implements CellStyle {

    private final XSSFWorkbook workbook;
    private final int xfId, styleXfId;
    private XSSFCellAlignment cellAlignment;
    private IndexedColorMap defaultIndexedColorMap = new DefaultIndexedColorMap();

    public XCellStyle(XSSFWorkbook workbook, int xfId, int styleXfId) {
        this.workbook = workbook;
        this.xfId = xfId;
        this.styleXfId = styleXfId;
    }

    private CTXf getCoreXf() {
        if(xfId < 0)
            return workbook.getStylesSource().getCellXfAt(0);
        else
            return workbook.getStylesSource().getCellXfAt(xfId);
    }

    private CTXf getStyleXf() {
        return workbook.getStylesSource().getCellStyleXfAt(styleXfId);
    }

    private CTXf getXf() {
        if(styleXfId > -1) return workbook.getStylesSource().getCellStyleXfAt(styleXfId);
        else return getCoreXf();
    }

    private CTBorder getCTBorder(){
        CTBorder ct;
        CTXf xf = getXf();
        if(xf.getApplyBorder()) {
            int idx = (int) xf.getBorderId();
            XSSFCellBorder cf = workbook.getStylesSource().getBorderAt(idx);
            ct = (CTBorder)cf.getCTBorder().copy();
        } else {
            ct = CTBorder.Factory.newInstance();
        }
        return ct;
    }

    private CTFill getCTFill(){
        CTFill ct;
        CTXf xf = getXf();
        if(xf.getApplyFill()) {
            int fillIndex = (int)xf.getFillId();
            XSSFCellFill cf = workbook.getStylesSource().getFillAt(fillIndex);

            ct = (CTFill)cf.getCTFill().copy();
        } else {
            ct = CTFill.Factory.newInstance();
        }
        return ct;
    }

    private XSSFCellAlignment getCellAlignment() {
        if (cellAlignment == null) {
            cellAlignment = new XSSFCellAlignment(getCTCellAlignment());
        }
        return cellAlignment;
    }

    private CTCellAlignment getCTCellAlignment() {
        CTXf xf = getXf();
        if (xf.getAlignment() == null) {
            xf.setAlignment(CTCellAlignment.Factory.newInstance());
        }
        return xf.getAlignment();
    }

    public void setBorderBottom(BorderStyle border) {
        CTBorder ct = getCTBorder();
        CTBorderPr pr = ct.isSetBottom() ? ct.getBottom() : ct.addNewBottom();
        if(border == BorderStyle.NONE) ct.unsetBottom();
        else pr.setStyle(STBorderStyle.Enum.forInt(border.getCode() + 1));

        int idx = workbook.getStylesSource().putBorder(new XSSFCellBorder(ct));

        CTXf xf = getXf();
        xf.setBorderId(idx);
        xf.setApplyBorder(true);
        getCoreXf().setBorderId(idx);
    }

    public void setBorderLeft(BorderStyle border) {
        CTBorder ct = getCTBorder();
        CTBorderPr pr = ct.isSetLeft() ? ct.getLeft() : ct.addNewLeft();
        if(border == BorderStyle.NONE) ct.unsetLeft();
        else pr.setStyle(STBorderStyle.Enum.forInt(border.getCode() + 1));

        int idx = workbook.getStylesSource().putBorder(new XSSFCellBorder(ct));

        CTXf xf = getXf();
        xf.setBorderId(idx);
        xf.setApplyBorder(true);
        getCoreXf().setBorderId(idx);
    }

    public void setBorderRight(BorderStyle border) {
        CTBorder ct = getCTBorder();
        CTBorderPr pr = ct.isSetRight() ? ct.getRight() : ct.addNewRight();
        if(border == BorderStyle.NONE) ct.unsetRight();
        else pr.setStyle(STBorderStyle.Enum.forInt(border.getCode() + 1));

        int idx = workbook.getStylesSource().putBorder(new XSSFCellBorder(ct));

        CTXf xf = getXf();
        xf.setBorderId(idx);
        xf.setApplyBorder(true);
        getCoreXf().setBorderId(idx);
    }

    public void setBorderTop(BorderStyle border) {
        CTBorder ct = getCTBorder();
        CTBorderPr pr = ct.isSetTop() ? ct.getTop() : ct.addNewTop();
        if(border == BorderStyle.NONE) ct.unsetTop();
        else pr.setStyle(STBorderStyle.Enum.forInt(border.getCode() + 1));

        int idx = workbook.getStylesSource().putBorder(new XSSFCellBorder(ct));

        CTXf xf = getXf();
        xf.setBorderId(idx);
        xf.setApplyBorder(true);
        getCoreXf().setBorderId(idx);
    }

    private void setBottomBorderColor(XSSFColor color) {
        CTBorder ct = getCTBorder();
        if(color == null && !ct.isSetBottom()) return;

        CTBorderPr pr = ct.isSetBottom() ? ct.getBottom() : ct.addNewBottom();
        if(color != null)  pr.setColor(color.getCTColor());
        else pr.unsetColor();

        int idx = workbook.getStylesSource().putBorder(new XSSFCellBorder(ct));

        CTXf xf = getXf();
        xf.setBorderId(idx);
        xf.setApplyBorder(true);
        getCoreXf().setBorderId(idx);
    }

    public void setBottomBorderColor(short color) {
        XSSFColor clr = new XSSFColor(defaultIndexedColorMap);
        clr.setIndexed(color);
        setBottomBorderColor(clr);
    }

    private void setLeftBorderColor(XSSFColor color) {
        CTBorder ct = getCTBorder();
        if(color == null && !ct.isSetLeft()) return;

        CTBorderPr pr = ct.isSetLeft() ? ct.getLeft() : ct.addNewLeft();
        if(color != null)  pr.setColor(color.getCTColor());
        else pr.unsetColor();

        int idx = workbook.getStylesSource().putBorder(new XSSFCellBorder(ct));

        CTXf xf = getXf();
        xf.setBorderId(idx);
        xf.setApplyBorder(true);
        getCoreXf().setBorderId(idx);
    }

    public void setLeftBorderColor(short color) {
        XSSFColor clr = new XSSFColor(defaultIndexedColorMap);
        clr.setIndexed(color);
        setLeftBorderColor(clr);
    }

    private void setRightBorderColor(XSSFColor color) {
        CTBorder ct = getCTBorder();
        if(color == null && !ct.isSetRight()) return;

        CTBorderPr pr = ct.isSetRight() ? ct.getRight() : ct.addNewRight();
        if(color != null)  pr.setColor(color.getCTColor());
        else pr.unsetColor();

        int idx = workbook.getStylesSource().putBorder(new XSSFCellBorder(ct));

        CTXf xf = getXf();
        xf.setBorderId(idx);
        xf.setApplyBorder(true);
        getCoreXf().setBorderId(idx);
    }

    public void setRightBorderColor(short color) {
        XSSFColor clr = new XSSFColor(defaultIndexedColorMap);
        clr.setIndexed(color);
        setRightBorderColor(clr);
    }

    private void setTopBorderColor(XSSFColor color) {
        CTBorder ct = getCTBorder();
        if(color == null && !ct.isSetTop()) return;

        CTBorderPr pr = ct.isSetTop() ? ct.getTop() : ct.addNewTop();
        if(color != null)  pr.setColor(color.getCTColor());
        else pr.unsetColor();

        int idx = workbook.getStylesSource().putBorder(new XSSFCellBorder(ct));

        CTXf xf = getXf();
        xf.setBorderId(idx);
        xf.setApplyBorder(true);
        getCoreXf().setBorderId(idx);
    }

    public void setTopBorderColor(short color) {
        XSSFColor clr = new XSSFColor(defaultIndexedColorMap);
        clr.setIndexed(color);
        setTopBorderColor(clr);
    }

    public void setDataFormat(String format) {
        DataFormat dataFormat = workbook.createDataFormat();
        short fmtId = dataFormat.getFormat(format);

        CTXf xf = getXf();
        xf.setApplyNumberFormat(true);
        xf.setNumFmtId(fmtId);
        getCoreXf().setNumFmtId(fmtId);
    }

    private void setFillBackgroundColor(XSSFColor color) {
        CTFill ct = getCTFill();
        CTPatternFill ptrn = ct.getPatternFill();
        if(color == null) {
            if(ptrn != null) ptrn.unsetBgColor();
        } else {
            if(ptrn == null) ptrn = ct.addNewPatternFill();
            ptrn.setBgColor(color.getCTColor());
        }

        StylesTable stylesTable = workbook.getStylesSource();
        int idx = stylesTable.putFill(new XSSFCellFill(ct, stylesTable.getIndexedColors()));

        CTXf xf = getXf();
        xf.setFillId(idx);
        xf.setApplyFill(true);
        getCoreXf().setFillId(idx);
    }

    public void setFillBackgroundColor(short bg) {
        XSSFColor clr = new XSSFColor(defaultIndexedColorMap);
        clr.setIndexed(bg);
        setFillBackgroundColor(clr);
    }

    private void setFillForegroundColor(XSSFColor color) {
        CTFill ct = getCTFill();

        CTPatternFill ptrn = ct.getPatternFill();
        if(color == null) {
            if(ptrn != null) ptrn.unsetFgColor();
        } else {
            if(ptrn == null) ptrn = ct.addNewPatternFill();
            ptrn.setFgColor(color.getCTColor());
        }

        StylesTable stylesTable = workbook.getStylesSource();
        int idx = stylesTable.putFill(new XSSFCellFill(ct, stylesTable.getIndexedColors()));

        CTXf xf = getXf();
        xf.setFillId(idx);
        xf.setApplyFill(true);
        getCoreXf().setFillId(idx);
    }

    public void setFillForegroundColor(short fg) {
        XSSFColor clr = new XSSFColor(defaultIndexedColorMap);
        clr.setIndexed(fg);
        setFillForegroundColor(clr);
    }

    public void setFillPattern(FillPatternType fp) {
        CTFill ct = getCTFill();
        CTPatternFill ptrn = ct.isSetPatternFill() ? ct.getPatternFill() : ct.addNewPatternFill();
        if(fp == FillPatternType.NO_FILL && ptrn.isSetPatternType()) ptrn.unsetPatternType();
        else ptrn.setPatternType(STPatternType.Enum.forInt(fp.getCode() + 1));

        StylesTable stylesTable = workbook.getStylesSource();
        int idx = stylesTable.putFill(new XSSFCellFill(ct, stylesTable.getIndexedColors()));

        CTXf xf = getXf();
        xf.setFillId(idx);
        xf.setApplyFill(true);
        getCoreXf().setFillId(idx);
    }

    public void setWrapText(boolean wrap) {
        getCellAlignment().setWrapText(wrap);
    }

    public static XCellStyle create(XSSFWorkbook workbook, String name) {

        int styleXfSize = 0;

        CTXf xf = CTXf.Factory.newInstance();
        xf.setNumFmtId(0);
        xf.setFontId(0);
        xf.setFillId(0);
        xf.setBorderId(0);

        if(name != null) {
            CTCellStyles ctCellStyles = workbook.getStylesSource().getCTStylesheet().getCellStyles();
            if(ctCellStyles == null) {
                ctCellStyles = workbook.getStylesSource().getCTStylesheet().addNewCellStyles();
                ctCellStyles.setCount(0);
            }
            if(ctCellStyles.getCount() == 0) {
                CTCellStyle standardCellStyle = ctCellStyles.addNewCellStyle();
                standardCellStyle.setName("Standard");
                standardCellStyle.setXfId(0);
                standardCellStyle.setBuiltinId(0);
                ctCellStyles.setCount(1);
            }

            CTXf styleXf = CTXf.Factory.newInstance();
            styleXf.setNumFmtId(0);
            styleXf.setFontId(0);
            styleXf.setFillId(0);
            styleXf.setBorderId(0);

            styleXfSize = workbook.getStylesSource().putCellStyleXf(styleXf);
            xf.setXfId(styleXfSize - 1);
            
            CTCellStyle ctCellStyle = ctCellStyles.addNewCellStyle();
            ctCellStyle.setName(name);
            ctCellStyle.setXfId(styleXfSize - 1);
            ctCellStyles.setCount(ctCellStyles.getCount() + 1);
        }

        int xfSize = workbook.getStylesSource().putCellXf(xf);

        return new XCellStyle(workbook, xfSize - 1, styleXfSize - 1);
    }

    /**
     * Used for querying user-named cell styles (style xf only).
     *
     * @param workbook
     * @param name
     * @return          XCellStyle with corresponding (named) style xf
     */
    public static XCellStyle get(XSSFWorkbook workbook, String name) {
        StylesTable stylesSource = workbook.getStylesSource();
        CTCellStyles ctCellStyles = stylesSource.getCTStylesheet().getCellStyles();

        if(ctCellStyles != null) {
            for(int i = 0; i < ctCellStyles.getCount(); i++) {
                CTCellStyle ctCellStyle = ctCellStyles.getCellStyleArray(i);
                if(ctCellStyle.getName().equals(name)) {
                    int styleXfId = (int) ctCellStyle.getXfId();

                    return new XCellStyle(workbook, -1, styleXfId);
                }
            }
        }

        return null;
    }

    public static void set(XSSFCell c, XCellStyle cs) {
        if(cs.xfId < 0) {
            // Only the style xf is of interest

            CTXf styleXf = cs.getStyleXf();

            CTXf xf = CTXf.Factory.newInstance();
            xf.setNumFmtId(styleXf.getNumFmtId());
            xf.setFontId(styleXf.getFontId());
            xf.setFillId(styleXf.getFillId());
            xf.setBorderId(styleXf.getBorderId());
            xf.setAlignment(styleXf.getAlignment());
            xf.setXfId(cs.styleXfId);

            int xfSize = cs.workbook.getStylesSource().putCellXf(xf);
            c.setCellStyle(new XSSFCellStyle(xfSize - 1, cs.styleXfId,
                    cs.workbook.getStylesSource(), cs.workbook.getTheme()));
        } else if(cs.styleXfId < 0) {
            // It's an unnamed cell style - only the core xf is of interest
            
            int styleXfId = 0;
            int id = (int) cs.workbook.getStylesSource().getCellXfAt(cs.xfId).getXfId();
            if(id > 0) styleXfId = id;
            
            c.setCellStyle(new XSSFCellStyle(cs.xfId, styleXfId, cs.workbook.getStylesSource(),
                    cs.workbook.getTheme()));
        } else
            c.setCellStyle(new XSSFCellStyle(cs.xfId, cs.styleXfId,
                    cs.workbook.getStylesSource(), cs.workbook.getTheme()));
    }
}
