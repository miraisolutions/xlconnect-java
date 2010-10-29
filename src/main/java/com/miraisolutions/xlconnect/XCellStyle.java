/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package com.miraisolutions.xlconnect;

import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
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
 * Partially uses code from XSSFCellStyle from
 * http://svn.apache.org/repos/asf/poi/trunk/src/ooxml/java/org/apache/poi/xssf/usermodel/XSSFCellStyle.java
 * 
 * @author Martin Studer, Mirai Solutions GmbH
 */
public class XCellStyle implements CellStyle {

    /**
     * getCoreXf().setXXX(...)
     * getStyleXf().setXXX(...)
     * getStyleXf().applyXXX(...)
     */

    private final XSSFWorkbook workbook;
    // private final StylesTable stylesSource;
    private final int xfId, styleXfId;
    // private final ThemesTable themes;
    private XSSFCellAlignment cellAlignment;

    public XCellStyle(XSSFWorkbook workbook, int xfId, int styleXfId) {
        this.workbook = workbook;
        this.xfId = xfId;
        this.styleXfId = styleXfId;
    }

    private CTXf getCoreXf() {
        return workbook.getStylesSource().getCellXfAt(xfId);
    }

    private CTXf getStyleXf() {
        return workbook.getStylesSource().getCellStyleXfAt(styleXfId);
    }

    private CTXf getXf() {
        if(styleXfId > -1) return workbook.getStylesSource().getCellStyleXfAt(styleXfId);
        else return workbook.getStylesSource().getCellXfAt(xfId);
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

    public void setBorderBottom(short border) {
        CTBorder ct = getCTBorder();
        CTBorderPr pr = ct.isSetBottom() ? ct.getBottom() : ct.addNewBottom();
        if(border == org.apache.poi.ss.usermodel.CellStyle.BORDER_NONE) ct.unsetBottom();
        else pr.setStyle(STBorderStyle.Enum.forInt(border + 1));

        int idx = workbook.getStylesSource().putBorder(new XSSFCellBorder(ct));

        CTXf xf = getXf();
        xf.setBorderId(idx);
        xf.setApplyBorder(true);
        getCoreXf().setBorderId(idx);
    }

    public void setDataFormat(String format) {
        DataFormat dataFormat = workbook.createDataFormat();
        short fmtId = dataFormat.getFormat(format);

        CTXf xf = getXf();
        xf.setApplyNumberFormat(true);
        xf.setNumFmtId(fmtId);
        getCoreXf().setNumFmtId(fmtId);
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

        int idx = workbook.getStylesSource().putFill(new XSSFCellFill(ct));

        CTXf xf = getXf();
        xf.setFillId(idx);
        xf.setApplyFill(true);
        getCoreXf().setFillId(idx);
    }

    public void setFillForegroundColor(short fg) {
        XSSFColor clr = new XSSFColor();
        clr.setIndexed(fg);
        setFillForegroundColor(clr);
    }

    public void setFillPattern(short fp) {
        CTFill ct = getCTFill();
        CTPatternFill ptrn = ct.isSetPatternFill() ? ct.getPatternFill() : ct.addNewPatternFill();
        if(fp == org.apache.poi.ss.usermodel.CellStyle.NO_FILL && ptrn.isSetPatternType()) ptrn.unsetPatternType();
        else ptrn.setPatternType(STPatternType.Enum.forInt(fp + 1));

        int idx = workbook.getStylesSource().putFill(new XSSFCellFill(ct));

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
     * @return          XCellStyle with core xf id = 0 and id of
     *                  corresponding (named) style xf
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
            xf.setXfId(cs.styleXfId);

            int xfSize = cs.workbook.getStylesSource().putCellXf(xf);
            c.setCellStyle(new XSSFCellStyle(xfSize - 1, cs.styleXfId,
                    cs.workbook.getStylesSource(), cs.workbook.getTheme()));
        } else if(cs.styleXfId < 0) {
            // It's an unnamed cell style - only the core xf is of interest

            c.setCellStyle(new XSSFCellStyle(cs.xfId, 0, cs.workbook.getStylesSource(),
                    cs.workbook.getTheme()));
        } else
            c.setCellStyle(new XSSFCellStyle(cs.xfId, cs.styleXfId,
                    cs.workbook.getStylesSource(), cs.workbook.getTheme()));
    }
}
