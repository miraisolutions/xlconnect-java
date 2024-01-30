/*
 *
    XLConnect
    Copyright (C) 2010-2024 Mirai Solutions GmbH

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

import com.miraisolutions.xlconnect.*;
import com.miraisolutions.xlconnect.data.DataFrame;
import com.miraisolutions.xlconnect.data.DataType;
import com.miraisolutions.xlconnect.data.ReadStrategy;
import com.miraisolutions.xlconnect.utils.RepeatableIterableUtils;
import com.miraisolutions.xlconnect.utils.SequenceLengthEncoding;
import com.miraisolutions.xlconnect.utils.SimpleSequence;
import org.apache.poi.common.usermodel.HyperlinkType;

import java.io.IOException;
import java.util.Arrays;

public final class RWorkbookWrapper {

    private final Workbook workbook;

    public RWorkbookWrapper(String filename, String password, boolean create) throws IOException {
        this.workbook = Workbook.getWorkbook(filename, password, create);
    }

    public RWorkbookWrapper(String filename, boolean create) throws IOException {
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

    public String[] getDefinedNames(boolean validOnly, String worksheetScope) {
        return workbook.getDefinedNames(validOnly, worksheetScope);
    }

    public void createSheet(String name) {
        workbook.createSheet(name);
    }

    public void createName(String name, String formula, boolean overwrite, String worksheetScope) {
        workbook.createName(name, worksheetScope, formula, overwrite);
    }

    public void removeName(String name, String worksheetScope) {
        workbook.removeName(name, worksheetScope);
    }

    public String getReferenceFormula(String name, String worksheetScope) {
        return workbook.getReferenceFormula(name, worksheetScope);
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

    public void writeNamedRegion(RDataFrameWrapper dataFrame, String name, boolean header, boolean overwriteFormulaCells, String worksheetScope) {
        workbook.writeNamedRegion(dataFrame.dataFrame, name, header, overwriteFormulaCells, worksheetScope);
    }

    private static DataType[] dataTypeFromString(String[] colTypes) {
        DataType[] ctypes = null;
        if (colTypes != null) {
            ctypes = Arrays.stream(colTypes)
                    .map(RWorkbookWrapper::getDataTypeFromString)
                    .toArray(DataType[]::new);
        }
        return ctypes;
    }

    public RDataFrameWithAttributes readNamedRegion(String name, boolean header, String[] colTypes,
                                             boolean forceConversion, String dateTimeFormat, boolean takeCached, int[] subset,
                                             String readStrategy, String worksheetScope) {
        DataFrameWithAttributes dataFrame = workbook.readNamedRegion(name, worksheetScope, header, ReadStrategy.valueOf(readStrategy.toUpperCase()),
                dataTypeFromString(colTypes), forceConversion, dateTimeFormat, takeCached, subset);
        return new RDataFrameWithAttributes(dataFrame.getAttributes(), new RDataFrameWrapper(dataFrame.getValue()));
    }

    public RDataFrameWrapper readTable(int worksheetIndex, String tableName, boolean header, String[] colTypes,
                                       boolean forceConversion, String dateTimeFormat, boolean takeCached, int[] subset, String readStrategy) {
        DataFrame dataFrame = workbook.readTable(worksheetIndex, tableName, header, ReadStrategy.valueOf(readStrategy.toUpperCase()),
                dataTypeFromString(colTypes), forceConversion, dateTimeFormat, takeCached, subset);
        return new RDataFrameWrapper(dataFrame);
    }

    public RDataFrameWrapper readTable(String worksheetName, String tableName, boolean header, String[] colTypes,
                                       boolean forceConversion, String dateTimeFormat, boolean takeCached, int[] subset, String readStrategy) {
        DataFrame dataFrame = workbook.readTable(worksheetName, tableName, header, ReadStrategy.valueOf(readStrategy.toUpperCase()),
                dataTypeFromString(colTypes), forceConversion, dateTimeFormat, takeCached, subset);
        return new RDataFrameWrapper(dataFrame);
    }

    public BooleanWithAttributes existsName(String name, String worksheetScope) {
        return workbook.existsName(name, worksheetScope);
    }

    public boolean existsSheet(String name) {
        return workbook.existsSheet(name);
    }

    public RDataFrameWrapper readWorksheet(int worksheetIndex, int startRow, int startCol, int endRow, int endCol,
                                           boolean header, String[] colTypes, boolean forceConversion, String dateTimeFormat, boolean takeCached,
                                           int[] subset, String readStrategy) {
        DataFrame dataFrame = workbook.readWorksheet(worksheetIndex, startRow, startCol, endRow, endCol, header,
                ReadStrategy.valueOf(readStrategy.toUpperCase()), dataTypeFromString(colTypes), forceConversion, dateTimeFormat,
                takeCached, subset, false, false);
        return new RDataFrameWrapper(dataFrame);
    }

    public RDataFrameWrapper readWorksheet(String worksheet, int startRow, int startCol, int endRow, int endCol,
                                           boolean header, String[] colTypes, boolean forceConversion, String dateTimeFormat, boolean takeCached,
                                           int[] subset, String readStrategy) {
        DataFrame dataFrame = workbook.readWorksheet(worksheet, startRow, startCol, endRow, endCol, header,
                ReadStrategy.valueOf(readStrategy.toUpperCase()), dataTypeFromString(colTypes), forceConversion, dateTimeFormat,
                takeCached, subset, false, false);
        return new RDataFrameWrapper(dataFrame);
    }

    public void writeWorksheet(RDataFrameWrapper dataFrame, int worksheetIndex, int startRow, int startCol, boolean header, boolean overwriteFormulaCells) {
        workbook.writeWorksheet(dataFrame.dataFrame, worksheetIndex, startRow, startCol, header, overwriteFormulaCells);
    }

    public void writeWorksheet(RDataFrameWrapper dataFrame, String worksheetName, int startRow, int startCol, boolean header, boolean overwriteFormulaCells) {
        workbook.writeWorksheet(dataFrame.dataFrame, worksheetName, startRow, startCol, header, overwriteFormulaCells);
    }

    public void writeWorksheet(RDataFrameWrapper dataFrame, int worksheetIndex, boolean header, boolean overwriteFormulaCells) {
        workbook.writeWorksheet(dataFrame.dataFrame, worksheetIndex, header, overwriteFormulaCells);
    }

    public void writeWorksheet(RDataFrameWrapper dataFrame, String worksheetName, boolean header, boolean overwriteFormulaCells) {
        workbook.writeWorksheet(dataFrame.dataFrame, worksheetName, header, overwriteFormulaCells);
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

    public void addImage(String filename, String name, boolean originalSize, String worksheetScope)
            throws IOException {
        workbook.addImage(filename, name, worksheetScope, originalSize);
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
        if (cs != null) {
            return new RCellStyleWrapper(cs);
        }
        throw new IllegalArgumentException("Cell style " + name + " does not exist");
    }

    public boolean existsCellStyle(String name) {
        return workbook.existsCellStyle(name);
    }

    public void setMissingValue(Object[] value) {
        workbook.setMissingValue(value);
    }

    private static DataType getDataTypeFromString(String dataType) {
        if ("BOOLEAN".equals(dataType))
            return DataType.Boolean;
        else if ("NUMERIC".equals(dataType))
            return DataType.Numeric;
        else if ("STRING".equals(dataType))
            return DataType.String;
        else if ("DATETIME".equals(dataType))
            return DataType.DateTime;
        else
            throw new IllegalArgumentException("Provided data type is not a valid data type!");
    }

    public void setDataFormat(String dataType, String format) {
        workbook.setDataFormat(getDataTypeFromString(dataType), format);
    }

    public void setCellStyleForDataType(String dataType, RCellStyleWrapper cellStyle) {
        workbook.setCellStyleForDataType(getDataTypeFromString(dataType), cellStyle.cellStyle);
    }

    public RCellStyleWrapper getCellStyleForDataType(String dataType) {
        return new RCellStyleWrapper(workbook.getCellStyleForDataType(getDataTypeFromString(dataType)));
    }

    public void setStyleAction(String action) {
        if ("XLCONNECT".equals(action))
            workbook.setStyleAction(StyleAction.XLCONNECT);
        else if ("DATATYPE".equals(action))
            workbook.setStyleAction(StyleAction.DATATYPE);
        else if ("NONE".equals(action))
            workbook.setStyleAction(StyleAction.NONE);
        else if ("PREDEFINED".equals(action))
            workbook.setStyleAction(StyleAction.PREDEFINED);
        else if ("STYLE_NAME_PREFIX".equals(action))
            workbook.setStyleAction(StyleAction.STYLE_NAME_PREFIX);
        else if ("DATA_FORMAT_ONLY".equals(action))
            workbook.setStyleAction(StyleAction.DATA_FORMAT_ONLY);
        else
            throw new IllegalArgumentException("Provided action is not a valid style action!");
    }

    public void setStyleNamePrefix(String prefix) {
        workbook.setStyleNamePrefix(prefix);
    }

    public void setCellStyleSheetName(SimpleSequence<String> sheetName, SequenceLengthEncoding row,
                                      SequenceLengthEncoding col, SimpleSequence<RCellStyleWrapper> cellStyle) {
        RepeatableIterableUtils.foreach(sheetName, row, col, cellStyle,
                (sheet, row1, col1, cellStyle1) -> workbook.setCellStyle(sheet, row1, col1, cellStyle1.cellStyle)
        );
    }

    public void setCellStyleSheetIndex(SimpleSequence<Integer> sheetIndex, SequenceLengthEncoding row,
                                       SequenceLengthEncoding col, SimpleSequence<RCellStyleWrapper> cellStyle) {
        RepeatableIterableUtils.foreach(sheetIndex, row, col, cellStyle,
                (sheet, row1, col1, cellStyle1) -> workbook.setCellStyle(sheet, row1, col1, cellStyle1.cellStyle)
        );
    }

    public void setCellStyleFormula(String formula, RCellStyleWrapper cellStyle) {
        workbook.setCellStyle(formula, cellStyle.cellStyle);
    }

    public void setHyperlinkSheetIndex(SimpleSequence<Integer> sheetIndex, SequenceLengthEncoding row,
                                       SequenceLengthEncoding col, SimpleSequence<String> type, SimpleSequence<String> address) {
        RepeatableIterableUtils.foreach(sheetIndex, row, col, type, address,
                (sheet, row1, col1, type1, address1) -> workbook.setHyperlink(sheet, row1, col1, HyperlinkType.valueOf(type1), address1)
        );
    }

    public void setHyperlinkSheetName(SimpleSequence<String> sheetName, SequenceLengthEncoding row,
                                      SequenceLengthEncoding col, SimpleSequence<String> type, SimpleSequence<String> address) {
        RepeatableIterableUtils.foreach(sheetName, row, col, type, address,
                (sheet, row1, col1, type1, address1) -> workbook.setHyperlink(sheet, row1, col1, HyperlinkType.valueOf(type1), address1)
        );
    }

    public void setHyperlinkFormula(String formula, String type, String address) {
        workbook.setHyperlink(formula, HyperlinkType.valueOf(type), address);
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
        return workbook.getAndClearWarnings().toArray(new String[0]);
    }

    public void onErrorCell(String behavior) {
        if ("STOP".equals(behavior))
            workbook.onErrorCell(ErrorBehavior.THROW_EXCEPTION);
        else
            workbook.onErrorCell(ErrorBehavior.WARN);
    }

    public void save(String file) throws IOException {
        workbook.save(file);
    }

    public void save() throws IOException {
        workbook.save();
    }

    public void setCellFormula(String sheetName, int row, int col, String formula) {
        workbook.setCellFormula(sheetName, row, col, formula);
    }

    public void setCellFormula(int sheetIndex, int row, int col, String formula) {
        workbook.setCellFormula(sheetIndex, row, col, formula);
    }

    public String getCellFormula(int sheetIndex, int row, int col) {
        return workbook.getCellFormula(sheetIndex, row, col);
    }

    public String getCellFormula(String sheetName, int row, int col) {
        return workbook.getCellFormula(sheetName, row, col);
    }

    public int[] getReferenceCoordinates(String name) {
        return workbook.getReferenceCoordinates(name);
    }

    public int[] getReferenceCoordinatesForName(String name, String worksheetScope) {
        return workbook.getReferenceCoordinatesForName(name, worksheetScope);
    }

    public int[] getReferenceCoordinatesForTable(int sheetIndex, String tableName) {
        return workbook.getReferenceCoordinatesForTable(sheetIndex, tableName);
    }

    public int[] getReferenceCoordinatesForTable(String sheetName, String tableName) {
        return workbook.getReferenceCoordinatesForTable(sheetName, tableName);
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
        return Math.max(workbook.getLastRow(sheetIndex), 0);
    }

    public int getLastRow(String sheetName) {
        return Math.max(workbook.getLastRow(sheetName), 0);
    }

    public int getLastColumn(int sheetIndex) {
        return workbook.getLastColumn(sheetIndex);
    }

    public int getLastColumn(String sheetName) {
        return workbook.getLastColumn(sheetName);
    }


    public void appendNamedRegion(RDataFrameWrapper data, String name, boolean header, boolean overwriteFormulaCells, String worksheetScope) {
        workbook.appendNamedRegion(data.dataFrame, name, worksheetScope, header, overwriteFormulaCells);
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

    public void clearRange(int sheetIndex, int[] coords) {
        workbook.clearRange(sheetIndex, coords);
    }

    public void clearRange(String sheetName, int[] coords) {
        workbook.clearRange(sheetName, coords);
    }

    public void clearRangeFromReference(String reference) {
        workbook.clearRangeFromReference(reference);
    }

    public void clearNamedRegion(String name, String worksheetScope) {
        workbook.clearNamedRegion(name, worksheetScope);
    }

    public void createFreezePane(int sheetIndex, int colSplit, int rowSplit, int leftColumn, int topRow) {
        workbook.createFreezePane(sheetIndex, colSplit, rowSplit, leftColumn, topRow);
    }

    public void createFreezePane(String sheetName, int colSplit, int rowSplit, int leftColumn, int topRow) {
        workbook.createFreezePane(sheetName, colSplit, rowSplit, leftColumn, topRow);
    }

    public void createFreezePane(int sheetIndex, int colSplit, int rowSplit) {
        workbook.createFreezePane(sheetIndex, colSplit, rowSplit);
    }

    public void createFreezePane(String sheetName, int colSplit, int rowSplit) {
        workbook.createFreezePane(sheetName, colSplit, rowSplit);
    }

    public void createSplitPane(int sheetIndex, int xSplitPos, int ySplitPos, int leftColumn, int topRow) {
        workbook.createSplitPane(sheetIndex, xSplitPos, ySplitPos, leftColumn, topRow);
    }

    public void createSplitPane(String sheetName, int xSplitPos, int ySplitPos, int leftColumn, int topRow) {
        workbook.createSplitPane(sheetName, xSplitPos, ySplitPos, leftColumn, topRow);
    }

    public void removePane(int sheetIndex) {
        workbook.removePane(sheetIndex);
    }

    public void removePane(String sheetName) {
        workbook.removePane(sheetName);
    }

    public void setSheetColor(int sheetIndex, int color) {
        workbook.setSheetColor(sheetIndex, color);
    }

    public void setSheetColor(String sheetName, int color) {
        workbook.setSheetColor(sheetName, color);
    }

    public int[] getBoundingBox(int sheetIndex, int startRow, int startCol, int endRow, int endCol,
                                boolean autofitRow, boolean autofitCol) {
        return workbook.getBoundingBox(sheetIndex, startRow, startCol, endRow, endCol,
                autofitRow, autofitCol);
    }

    public int[] getBoundingBox(String sheetName, int startRow, int startCol, int endRow, int endCol,
                                boolean autofitRow, boolean autofitCol) {
        return workbook.getBoundingBox(sheetName, startRow, startCol, endRow, endCol,
                autofitRow, autofitCol);
    }

    public String[] getTables(int sheetIndex) {
        return workbook.getTables(sheetIndex);
    }

    public String[] getTables(String sheetName) {
        return workbook.getTables(sheetName);
    }
}
