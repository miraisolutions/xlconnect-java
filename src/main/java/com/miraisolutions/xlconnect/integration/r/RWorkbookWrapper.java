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

    public void save() throws FileNotFoundException, IOException {
        workbook.save();
    }
}
