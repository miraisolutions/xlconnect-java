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

package com.miraisolutions.xlconnect.data;

import com.miraisolutions.xlconnect.ErrorBehavior;
import com.miraisolutions.xlconnect.utils.CellUtils;
import java.util.ArrayList;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;


public class DefaultColumnBuilder extends ColumnBuilder {
    
    // The following split is done for performance reasons
    protected String[] missingValueStrings;
    protected double[] missingValueNumbers;
    
    public DefaultColumnBuilder(int nrows, boolean forceConversion,
            boolean takeCached, FormulaEvaluator evaluator, ErrorBehavior onErrorCell,
            Object[] missingValue, String dateTimeFormat) {
        
        super(nrows, forceConversion, takeCached, evaluator, onErrorCell, dateTimeFormat);
        
        // Split missing values into missing values for strings and doubles
        // (for better performance later on)
        ArrayList<Double> missingNum = new ArrayList<Double>();
        ArrayList<String> missingStr = new ArrayList<String>();
        for(int i = 0; i < missingValue.length; i++) {
            if(missingValue[i] instanceof String) {
                missingStr.add((String) missingValue[i]);
            } else if(missingValue[i] instanceof Double) {
                missingNum.add((Double) missingValue[i]);
            }
        }
        
        missingValueStrings = missingStr.toArray(new String[missingStr.size()]);
        missingValueNumbers = new double[missingNum.size()];
        for(int i = 0; i < missingNum.size(); i++) {
            missingValueNumbers[i] = missingNum.get(i).doubleValue();
        }
    }
    
    @Override
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
                if(DateUtil.isCellDateFormatted(c))
                    addValue(c, cv, DataType.DateTime);
                else {
                    boolean missing = false;
                    for(int i = 0; i < missingValueNumbers.length; i++) {
                        if(cv.getNumberValue() == missingValueNumbers[i]) {
                            missing = true;
                            break;
                        }
                    }
                    if(missing)
                        addMissing();
                    else
                        addValue(c, cv, DataType.Numeric);
                }
                break;
            case STRING:
                boolean missing = false;
                for(int i = 0; i < missingValueStrings.length; i++) {
                    if(cv.getStringValue() == null || cv.getStringValue().equals(missingValueStrings[i])) {
                        missing = true;
                        break;
                    }
                }
                if(missing)
                    addMissing();
                else
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
