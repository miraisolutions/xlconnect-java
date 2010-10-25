/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package com.miraisolutions.xlconnect;

import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.model.ThemesTable;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellAlignment;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorder;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorderPr;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCell;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCellAlignment;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCellStyle;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCellStyleXfs;
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

    private final StylesTable stylesSource;
    private final CTCellStyle cellStyle;
    private final int xfId, styleXfId;
    private final ThemesTable themes;
    private XSSFCellAlignment cellAlignment;

    public XCellStyle(StylesTable stylesSource, CTCellStyle cellStyle, int xfId, int styleXfId,
            ThemesTable themes) {
        this.stylesSource = stylesSource;
        this.xfId = xfId;
        this.styleXfId = styleXfId;
        this.cellStyle = cellStyle;
        this.themes = themes;
    }

    public org.apache.poi.ss.usermodel.CellStyle getPOICellStyle() {
        return new XSSFCellStyle(xfId, styleXfId, stylesSource, themes);
    }

    private CTXf getCoreXf() {
        return stylesSource.getCellXfAt(xfId);
    }

    private CTXf getStyleXf() {
        return stylesSource.getCellStyleXfAt(styleXfId);
    }

    private CTBorder getCTBorder(){
        CTBorder ct;
        CTXf styleXf = getStyleXf();
        if(styleXf.getApplyBorder()) {
        // if(getCoreXf().getApplyBorder()) {
            int idx = (int) styleXf.getBorderId();
            XSSFCellBorder cf = stylesSource.getBorderAt(idx);
            ct = (CTBorder)cf.getCTBorder().copy();
        } else {
            ct = CTBorder.Factory.newInstance();
        }
        return ct;
    }

    private CTFill getCTFill(){
        CTFill ct;
        CTXf styleXf = getStyleXf();
        if(styleXf.getApplyFill()) {
        // if(getCoreXf().getApplyFill()) {
            int fillIndex = (int)styleXf.getFillId();
            XSSFCellFill cf = stylesSource.getFillAt(fillIndex);

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
        CTXf styleXf = getStyleXf();
        if (styleXf.getAlignment() == null) {
            styleXf.setAlignment(CTCellAlignment.Factory.newInstance());
        }
        return styleXf.getAlignment();
    }

    public void setBorderBottom(short border) {
        CTBorder ct = getCTBorder();
        CTBorderPr pr = ct.isSetBottom() ? ct.getBottom() : ct.addNewBottom();
        if(border == org.apache.poi.ss.usermodel.CellStyle.BORDER_NONE) ct.unsetBottom();
        else pr.setStyle(STBorderStyle.Enum.forInt(border + 1));

        int idx = stylesSource.putBorder(new XSSFCellBorder(ct));

        CTXf styleXf = getStyleXf();
        // ???
        getCoreXf().setBorderId(idx);
        styleXf.setBorderId(idx);
        styleXf.setApplyBorder(true);
        // getCoreXf().setApplyBorder(true);
    }

    public void setDataFormat(short format) {
        CTXf styleXf = getStyleXf();
        styleXf.setApplyNumberFormat(true);
        // getCoreXf().setApplyNumberFormat(true);
        styleXf.setNumFmtId(format);
        // ???
        getCoreXf().setNumFmtId(format);
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

        int idx = stylesSource.putFill(new XSSFCellFill(ct));

        CTXf styleXf = getStyleXf();
        // ???
        getCoreXf().setFillId(idx);
        styleXf.setFillId(idx);
        styleXf.setApplyFill(true);
        // getCoreXf().setApplyFill(true);
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

        int idx = stylesSource.putFill(new XSSFCellFill(ct));

        CTXf styleXf = stylesSource.getCellStyleXfAt(styleXfId);
        // ???
        getCoreXf().setFillId(idx);
        styleXf.setFillId(idx);
        styleXf.setApplyFill(true);
        // getCoreXf().setApplyFill(true);
    }

    public void setWrapText(boolean wrap) {
        getCellAlignment().setWrapText(wrap);
    }

    public static XCellStyle create(XSSFWorkbook workbook, String name) {

//        CTCellStyleXfs ctCellStyleXfs = workbook.getStylesSource().getCTStylesheet().getCellStyleXfs();
//        if(ctCellStyleXfs == null) {
//            ctCellStyleXfs = workbook.getStylesSource().getCTStylesheet().addNewCellStyleXfs();
//            ctCellStyleXfs.setCount(0);
//        }
//        if(ctCellStyleXfs.getCount() == 0) {
//            CTXf standardXf = ctCellStyleXfs.addNewXf();
//            standardXf.setNumFmtId(0);
//            standardXf.setFontId(0);
//            standardXf.setFillId(0);
//            standardXf.setBorderId(0);
//            ctCellStyleXfs.setCount(1);
//        }

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

        int styleXfSize = workbook.getStylesSource().putCellStyleXf(styleXf);

        CTXf xf = CTXf.Factory.newInstance();
        xf.setNumFmtId(0);
        xf.setFontId(0);
        xf.setFillId(0);
        xf.setBorderId(0);
        xf.setXfId(styleXfSize - 1);
        int xfSize = workbook.getStylesSource().putCellXf(xf);

        CTCellStyle ctCellStyle = ctCellStyles.addNewCellStyle();
        ctCellStyle.setName(name);
        // ctCellStyle.setXfId(workbook.getStylesSource()._getStyleXfsSize() - 1);
        ctCellStyle.setXfId(styleXfSize - 1);
        ctCellStyles.setCount(ctCellStyles.getCount() + 1);

        return new XCellStyle(workbook.getStylesSource(), ctCellStyle, xfSize - 1,
                styleXfSize - 1, workbook.getTheme());
    }

    public static XCellStyle get(XSSFWorkbook workbook, String name) {
        StylesTable stylesSource = workbook.getStylesSource();
        CTCellStyles ctCellStyles = stylesSource.getCTStylesheet().getCellStyles();

        if(ctCellStyles != null) {
            for(int i = 0; i < ctCellStyles.getCount(); i++) {
                CTCellStyle ctCellStyle = ctCellStyles.getCellStyleArray(i);
                if(ctCellStyle.getName().equals(name)) {
                    int styleXfId = (int) ctCellStyle.getXfId();
                    
                    return new XCellStyle(stylesSource, ctCellStyle, stylesSource._getXfsSize() - 1,
                            styleXfId, workbook.getTheme());
                }
            }
        }

        return null;
    }
}
