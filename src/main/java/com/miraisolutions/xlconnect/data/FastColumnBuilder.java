/*
 *
    XLConnect
    Copyright (C) 2013-2018 Mirai Solutions GmbH

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
package com.miraisolutions.xlconnect.data;

import com.miraisolutions.xlconnect.ErrorBehavior;
import com.miraisolutions.xlconnect.utils.CellUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;

/**
 *
 * @author mstuder
 */
public class FastColumnBuilder extends ColumnBuilder {
    
    public FastColumnBuilder(int nrows, boolean forceConversion,
            boolean takeCached, FormulaEvaluator evaluator, ErrorBehavior onErrorCell,
            String dateTimeFormat) {
        
        super(nrows, forceConversion, takeCached, evaluator, onErrorCell, dateTimeFormat);
    }
    
    protected void handleCell(Cell c, CellValue cv) {
        String msg;
        // Determine (evaluated) cell data type
        switch(cv.getCellType()) {
            case BLANK:
                addMissing();
                return;
            case BOOLEAN:
                addValue(c, cv, DataType.Boolean);
                break;
            case NUMERIC:
                addValue(c, cv, DataType.Numeric);
                break;
            case STRING:
                addValue(c, cv, DataType.String);
                break;
            case FORMULA:
                msg = "Formula detected in already evaluated cell " + CellUtils.formatAsString(c) + "!";
                cellError(msg);
                break;
            case ERROR:
                msg = "Error detected in cell " + CellUtils.formatAsString(c) + " - " + CellUtils.getErrorMessage(cv.getErrorValue());
                cellError(msg);
                break;
            default:
                msg = "Unexpected cell type detected for cell " + CellUtils.formatAsString(c) + "!";
                cellError(msg);
        }
    }
}
