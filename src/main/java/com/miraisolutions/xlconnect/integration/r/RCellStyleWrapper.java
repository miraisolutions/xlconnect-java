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

package com.miraisolutions.xlconnect.integration.r;

import com.miraisolutions.xlconnect.CellStyle;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;

public class RCellStyleWrapper {

    final CellStyle cellStyle;

    public RCellStyleWrapper(CellStyle cellStyle) {
        this.cellStyle = cellStyle;
    }

    public void setBorder(String[] side, int[] border, int[] color) {
        assert border.length == side.length && color.length == side.length;

        for(int i = 0; i < side.length; i++) {
            BorderStyle bs = BorderStyle.valueOf((short) border[i]);
            short bc = (short) color[i];

            if("bottom".equals(side[i])) {
                cellStyle.setBorderBottom(bs);
                cellStyle.setBottomBorderColor(bc);
            } else if("left".equals(side[i])) {
                cellStyle.setBorderLeft(bs);
                cellStyle.setLeftBorderColor(bc);
            } else if("right".equals(side[i])) {
                cellStyle.setBorderRight(bs);
                cellStyle.setRightBorderColor(bc);
            } else if("top".equals(side[i])) {
                cellStyle.setBorderTop(bs);
                cellStyle.setTopBorderColor(bc);
            } else
                throw new IllegalArgumentException("Undefined border side: '" + side[i] + "'");
        }

    }

    public void setDataFormat(String format) {
        cellStyle.setDataFormat(format);
    }

    public void setFillBackgroundColor(int bg) {
        cellStyle.setFillBackgroundColor((short) bg);
    }

    public void setFillForegroundColor(int fp) {
        cellStyle.setFillForegroundColor((short) fp);
    }

    public void setFillPattern(int bg) { cellStyle.setFillPattern(FillPatternType.forInt(bg)); }

    public void setWrapText(boolean wrap) {
        cellStyle.setWrapText(wrap);
    }

    public String[] retrieveWarnings() {
        return cellStyle.retrieveWarnings();
    }
}
