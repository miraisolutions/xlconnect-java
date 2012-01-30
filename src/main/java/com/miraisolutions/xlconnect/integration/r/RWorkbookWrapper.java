/*
 *
    XLConnect
    Copyright (C) 2010 Mirai Solutions GmbH

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

package com.miraisolutions.xlconnect.integration.r;

import com.miraisolutions.xlconnect.CellStyle;
import com.miraisolutions.xlconnect.ErrorBehavior;
import com.miraisolutions.xlconnect.StyleAction;
import com.miraisolutions.xlconnect.Workbook;
import com.miraisolutions.xlconnect.data.DataFrame;
import com.miraisolutions.xlconnect.data.DataType;
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

    public int getSheetPos(String sheetName) {
        return workbook.getSheetPos(sheetName);
    }

    public void setSheetPos(String sheetName, int pos) {
        workbook.setSheetPos(sheetName, pos);
    }

    public String[] getDefinedNames(boolean validOnly) {
        return workbook.getDefinedNames(validOnly);
    }

    public void createSheet(String name) {
        workbook.createSheet(name);
    }

    public void createName(String name, String formula, boolean overwrite) {
        workbook.createName(name, formula, overwrite);
    }

    public void removeName(String name) {
        workbook.removeName(name);
    }

    public String getReferenceFormula(String name) {
        return workbook.getReferenceFormula(name);
    }

    public void removeSheet(String name) {
        workbook.removeSheet(name);
    }

    public void removeSheet(int sheetIndex) {
        workbook.removeSheet(sheetIndex);
    }

    public void renameSheet(String name, String newName) {
        workbook.renameSheet(name, newName);
    }

    public void renameSheet(int sheetIndex, String newName) {
        workbook.renameSheet(sheetIndex, newName);
    }

    public void cloneSheet(int index, String newName) {
        workbook.cloneSheet(index, newName);
    }

    public void cloneSheet(String name, String newName) {
        workbook.cloneSheet(name, newName);
    }

    public void writeNamedRegion(RDataFrameWrapper dataFrame, String name, boolean header) {
        workbook.writeNamedRegion(dataFrame.dataFrame, name, header);
    }

    private static DataType[] fromString(String[] colTypes) {
        DataType[] ctypes = null;
        if(colTypes != null) {
            ctypes = new DataType[colTypes.length];
            for(int i = 0; i < colTypes.length; i++)
                ctypes[i] = fromString(colTypes[i]);
        }
        return ctypes;
    }

    public RDataFrameWrapper readNamedRegion(String name, boolean header, String[] colTypes) {
        DataFrame dataFrame = workbook.readNamedRegion(name, header, fromString(colTypes));
        return new RDataFrameWrapper(dataFrame);
    }

    public boolean existsName(String name) {
        return workbook.existsName(name);
    }

    public boolean existsSheet(String name) {
        return workbook.existsSheet(name);
    }

    public RDataFrameWrapper readWorksheet(int worksheetIndex, boolean header, String[] colTypes) {
        DataFrame dataFrame = workbook.readWorksheet(worksheetIndex, header, fromString(colTypes));
        return new RDataFrameWrapper(dataFrame);
    }

    public RDataFrameWrapper readWorksheet(String worksheetName, boolean header, String[] colTypes) {
        DataFrame dataFrame = workbook.readWorksheet(worksheetName, header, fromString(colTypes));
        return new RDataFrameWrapper(dataFrame);
    }

    public RDataFrameWrapper readWorksheet(int worksheetIndex, int startRow, int startCol, int endRow, int endCol, 
            boolean header, String[] colTypes) {
        DataFrame dataFrame = workbook.readWorksheet(worksheetIndex, startRow, startCol, endRow, endCol, header,
                fromString(colTypes));
        return new RDataFrameWrapper(dataFrame);
    }

    public RDataFrameWrapper readWorksheet(String worksheet, int startRow, int startCol, int endRow, int endCol, 
            boolean header, String colTypes[]) {
        DataFrame dataFrame = workbook.readWorksheet(worksheet, startRow, startCol, endRow, endCol, header,
                fromString(colTypes));
        return new RDataFrameWrapper(dataFrame);
    }

    public void writeWorksheet(RDataFrameWrapper dataFrame, int worksheetIndex, int startRow, int startCol, boolean header) {
        workbook.writeWorksheet(dataFrame.dataFrame, worksheetIndex, startRow, startCol, header);
    }

    public void writeWorksheet(RDataFrameWrapper dataFrame, String worksheetName, int startRow, int startCol, boolean header) {
        workbook.writeWorksheet(dataFrame.dataFrame, worksheetName, startRow, startCol, header);
    }

    public void writeWorksheet(RDataFrameWrapper dataFrame, int worksheetIndex, boolean header) {
        workbook.writeWorksheet(dataFrame.dataFrame, worksheetIndex, header);
    }

    public void writeWorksheet(RDataFrameWrapper dataFrame, String worksheetName, boolean header) {
        workbook.writeWorksheet(dataFrame.dataFrame, worksheetName, header);
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

    public void addImage(String filename, String name, boolean originalSize)
            throws FileNotFoundException, IOException {
        workbook.addImage(filename, name, originalSize);
    }

    public RCellStyleWrapper createCellStyle(String name) {
        CellStyle cs = workbook.createCellStyle(name);
        return new RCellStyleWrapper(cs);
    }

    public RCellStyleWrapper createCellStyle() {
        CellStyle cs = workbook.createCellStyle();
        return new RCellStyleWrapper(cs);
    }

    public RCellStyleWrapper getCellStyle(String name) {
        CellStyle cs = workbook.getCellStyle(name);
        if(cs != null) {
            return new RCellStyleWrapper(cs);
        }
        throw new IllegalArgumentException("Cell style " + name + " does not exist");
    }

    public void setMissingValue(String[] value) {
        workbook.setMissingValue(value);
    }

    private static DataType fromString(String dataType) {
        if("BOOLEAN".equals(dataType))
            return DataType.Boolean;
        else if("NUMERIC".equals(dataType))
            return DataType.Numeric;
        else if("STRING".equals(dataType))
            return DataType.String;
        else if("DATETIME".equals(dataType))
            return DataType.DateTime;
        else
            throw new IllegalArgumentException("Provided data type is not a valid data type!");
    }

    public void setDataFormat(String dataType, String format) {
        workbook.setDataFormat(fromString(dataType), format);
    }
    
    public void setStyleAction(String action) {
        if("XLCONNECT".equals(action))
            workbook.setStyleAction(StyleAction.XLCONNECT);
        else if("NONE".equals(action))
            workbook.setStyleAction(StyleAction.NONE);
        else if("PREDEFINED".equals(action))
            workbook.setStyleAction(StyleAction.PREDEFINED);
        else if("STYLE_NAME_PREFIX".equals(action))
            workbook.setStyleAction(StyleAction.STYLE_NAME_PREFIX);
        else if("DATA_FORMAT_ONLY".equals(action))
            workbook.setStyleAction(StyleAction.DATA_FORMAT_ONLY);
        else
            throw new IllegalArgumentException("Provided action is not a valid style action!");
    }

    public void setStyleNamePrefix(String prefix) {
        workbook.setStyleNamePrefix(prefix);
    }

    public void setCellStyle(String sheetName, int row, int col, RCellStyleWrapper cellStyle) {
        workbook.setCellStyle(sheetName, row, col, cellStyle.cellStyle);
    }

    public void setCellStyle(int sheetIndex, int row, int col, RCellStyleWrapper cellStyle) {
        workbook.setCellStyle(sheetIndex, row, col, cellStyle.cellStyle);
    }

    public void setColumnWidth(int sheetIndex, int columnIndex, int width) {
        workbook.setColumnWidth(sheetIndex, columnIndex, width);
    }

    public void setColumnWidth(String sheetName, int columnIndex, int width) {
        workbook.setColumnWidth(sheetName, columnIndex, width);
    }

    public void setRowHeight(int sheetIndex, int rowIndex, double height) {
        workbook.setRowHeight(sheetIndex, rowIndex, (float) height);
    }

    public void setRowHeight(String sheetName, int rowIndex, double height) {
        workbook.setRowHeight(sheetName, rowIndex, (float) height);
    }

    public void mergeCells(int sheetIndex, String reference) {
        workbook.mergeCells(sheetIndex, reference);
    }

    public void mergeCells(String sheetName, String reference) {
        workbook.mergeCells(sheetName, reference);
    }

    public void unmergeCells(int sheetIndex, String reference) {
        workbook.unmergeCells(sheetIndex, reference);
    }

    public void unmergeCells(String sheetName, String reference) {
        workbook.unmergeCells(sheetName, reference);
    }

    public String[] retrieveWarnings() {
        return workbook.retrieveWarnings();
    }
    
    public void onErrorCell(String behavior) {
        if("STOP".equals(behavior))
            workbook.onErrorCell(ErrorBehavior.THROW_EXCEPTION);
        else
            workbook.onErrorCell(ErrorBehavior.WARN);
    }

    public void save(String file) throws FileNotFoundException, IOException {
        workbook.save(file);
    }

    public void save() throws FileNotFoundException, IOException {
        workbook.save();
    }

    public void setCellFormula(String sheetName, int row, int col, String formula) {
        workbook.setCellFormula(sheetName, row, col, formula);
    }

    public void setCellFormula(int sheetIndex, int row, int col, String formula) {
        workbook.setCellFormula(sheetIndex, row, col, formula);
    }

    public String getCellFormula(int sheetIndex, int row, int col) {
        return workbook.getCellFormula(sheetIndex,row,col);
    }

    public String getCellFormula(String sheetName, int row, int col) {
        return workbook.getCellFormula(sheetName,row,col);
    }

    public int[] getReferenceCoordinates(String name) {
	return workbook.getReferenceCoordinates(name);
    }

    public void setForceFormulaRecalculation(int sheetIndex, boolean value) {
        workbook.setForceFormulaRecalculation(sheetIndex, value);
    }

    public void setForceFormulaRecalculation(String sheetName, boolean value) {
        workbook.setForceFormulaRecalculation(sheetName, value);
    }

    public boolean getForceFormulaRecalculation(int sheetIndex) {
        return workbook.getForceFormulaRecalculation(sheetIndex);
    }

    public boolean getForceFormulaRecalculation(String sheetName) {
        return workbook.getForceFormulaRecalculation(sheetName);
    }

    public void setAutoFilter(int sheetIndex, String reference) {
        workbook.setAutoFilter(sheetIndex, reference);
    }

    public void setAutoFilter(String sheetName, String reference) {
        workbook.setAutoFilter(sheetName, reference);
    }

    public int getLastRow(int sheetIndex) {
        return workbook.getLastRow(sheetIndex);
    }

    public int getLastRow(String sheetName) {
        return workbook.getLastRow(sheetName);
    }

    public void appendNamedRegion(RDataFrameWrapper data, String name, boolean header) {
        workbook.appendNamedRegion(data.dataFrame, name, header);
    }

    public void appendWorksheet(RDataFrameWrapper data, int worksheetIndex, boolean header) {
        workbook.appendWorksheet(data.dataFrame, worksheetIndex, header);
    }

    public void appendWorksheet(RDataFrameWrapper data, String worksheetName, boolean header) {
        workbook.appendWorksheet(data.dataFrame, worksheetName, header);
    }
    
    public void clearSheet(int sheetIndex) {
        workbook.clearSheet(sheetIndex);
    }

    public void clearSheet(String sheetName) {
        workbook.clearSheet(sheetName);
    }

    public void setForceConversion(boolean forceConversion) {
        workbook.setForceConversion(forceConversion);
    }
}
