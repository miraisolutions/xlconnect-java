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

/**
 *
 * @author Martin Studer, Mirai Solutions GmbH
 */
public class RCellStyleWrapper {

    final CellStyle cellStyle;

    public RCellStyleWrapper(CellStyle cellStyle) {
        this.cellStyle = cellStyle;
    }

    public void setBorder(String[] side, int[] border, int[] color) {
        assert border.length == side.length && color.length == side.length;

        for(int i = 0; i < side.length; i++) {
            if("bottom".equals(side[i])) {
                cellStyle.setBorderBottom((short) border[i]);
                cellStyle.setBottomBorderColor((short) color[i]);
            } else if("left".equals(side[i])) {
                cellStyle.setBorderLeft((short) border[i]);
                cellStyle.setLeftBorderColor((short) color[i]);
            } else if("right".equals(side[i])) {
                cellStyle.setBorderRight((short) border[i]);
                cellStyle.setRightBorderColor((short) color[i]);
            } else if("top".equals(side[i])) {
                cellStyle.setBorderTop((short) border[i]);
                cellStyle.setTopBorderColor((short) color[i]);
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

    public void setFillPattern(int bg) {
        cellStyle.setFillPattern((short) bg);
    }

    public void setWrapText(boolean wrap) {
        cellStyle.setWrapText(wrap);
    }
}
