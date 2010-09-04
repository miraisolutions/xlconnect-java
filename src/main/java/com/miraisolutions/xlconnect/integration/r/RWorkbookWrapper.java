/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package com.miraisolutions.xlconnect.integration.r;

import com.miraisolutions.xlconnect.Workbook;
import com.miraisolutions.xlconnect.data.DataFrame;
import java.io.FileNotFoundException;
import java.io.IOException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

/**
 *
 * @author Martin Studer, Mirai Solutions GmbH
 */
public final class RWorkbookWrapper {

    private final Workbook workbook;

    public RWorkbookWrapper(String filename, boolean create) throws FileNotFoundException, IOException, InvalidFormatException {
        this.workbook = Workbook.getWorkbook(filename, create);
    }

    public String[] getSheets() {
        return workbook.getSheets();
    }

    public String[] getDefinedNames() {
        return workbook.getDefinedNames();
    }

    public void createSheet(String name) {
        workbook.createSheet(name);
    }

    public void removeName(String name) {
        workbook.removeName(name);
    }

    public void removeSheet(String name) {
        workbook.removeSheet(name);
    }

    public void writeNamedRegion(RDataFrameWrapper dataFrame, String name) {
        workbook.writeNamedRegion(dataFrame.dataFrame, name);
    }

    public void writeNamedRegion(RDataFrameWrapper dataFrame, String name, String location, boolean overwrite) {
        workbook.writeNamedRegion(dataFrame.dataFrame, name, location, overwrite);
    }

    public RDataFrameWrapper readNamedRegion(String name, boolean header) {
        DataFrame dataFrame = workbook.readNamedRegion(name, header);
        return new RDataFrameWrapper(dataFrame);
    }

    public boolean existsName(String name) {
        return workbook.existsName(name);
    }

    public boolean existsSheet(String name) {
        return workbook.existsSheet(name);
    }

    public RDataFrameWrapper readWorksheet(int worksheetIndex, boolean header) {
        DataFrame dataFrame = workbook.readWorksheet(worksheetIndex, header);
        return new RDataFrameWrapper(dataFrame);
    }

    public RDataFrameWrapper readWorksheet(String worksheetName, boolean header) {
        DataFrame dataFrame = workbook.readWorksheet(worksheetName, header);
        return new RDataFrameWrapper(dataFrame);
    }

    public RDataFrameWrapper readWorksheet(int worksheetIndex, int startRow, int startCol, int endRow, int endCol, boolean header) {
        DataFrame dataFrame = workbook.readWorksheet(worksheetIndex, startRow, startCol, endRow, endCol, header);
        return new RDataFrameWrapper(dataFrame);
    }

    public RDataFrameWrapper readWorksheet(String worksheet, int startRow, int startCol, int endRow, int endCol, boolean header) {
        DataFrame dataFrame = workbook.readWorksheet(worksheet, startRow, startCol, endRow, endCol, header);
        return new RDataFrameWrapper(dataFrame);
    }

    public void writeWorksheet(RDataFrameWrapper dataFrame, int worksheetIndex, int startRow, int startCol) {
        workbook.writeWorksheet(dataFrame.dataFrame, worksheetIndex, startRow, startCol);
    }

    public void writeWorksheet(RDataFrameWrapper dataFrame, String worksheetName, int startRow, int startCol, boolean create) {
        workbook.writeWorksheet(dataFrame.dataFrame, worksheetName, startRow, startCol, create);
    }

    public void writeWorksheet(RDataFrameWrapper dataFrame, int worksheetIndex) {
        workbook.writeWorksheet(dataFrame.dataFrame, worksheetIndex);
    }

    public void writeWorksheet(RDataFrameWrapper dataFrame, String worksheetName, boolean create) {
        workbook.writeWorksheet(dataFrame.dataFrame, worksheetName, create);
    }

    public int getActiveSheetIndex() {
        return workbook.getActiveSheetIndex();
    }

    public String getActiveSheetName() {
        return workbook.getActiveSheetName();
    }

    public void setActiveSheet(int sheetIndex) {
        workbook.setActiveSheet(sheetIndex);
    }

    public void setActiveSheet(String sheetName) {
        workbook.setActiveSheet(sheetName);
    }

    public void hideSheet(int sheetIndex, boolean veryHidden) {
        workbook.hideSheet(sheetIndex, veryHidden);
    }

    public void hideSheet(String sheetName, boolean veryHidden) {
        workbook.hideSheet(sheetName, veryHidden);
    }

    public void unhideSheet(int sheetIndex) {
        workbook.unhideSheet(sheetIndex);
    }

    public void unhideSheet(String sheetName) {
        workbook.unhideSheet(sheetName);
    }

    public boolean isSheetHidden(int sheetIndex) {
        return workbook.isSheetHidden(sheetIndex);
    }

    public boolean isSheetHidden(String sheetName) {
        return workbook.isSheetHidden(sheetName);
    }

    public boolean isSheetVeryHidden(int sheetIndex) {
        return workbook.isSheetVeryHidden(sheetIndex);
    }

    public boolean isSheetVeryHidden(String sheetName) {
        return workbook.isSheetVeryHidden(sheetName);
    }

    public void save() throws FileNotFoundException, IOException {
        workbook.save();
    }
}
