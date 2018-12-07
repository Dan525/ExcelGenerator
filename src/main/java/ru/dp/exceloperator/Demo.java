/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ru.dp.exceloperator;

import java.io.FileNotFoundException;
import java.io.IOException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

/**
 *
 * @author daniil_pozdeev
 */
public class Demo {
    public static void main(String[] args) throws IOException, FileNotFoundException, InvalidFormatException {
        ExcelScanner reader = new ExcelScanner("D:\\MyWorkspace\\protocol.xlsx");
        reader.scanSheet(0);
        ExcelUpdater updater = new ExcelUpdater(reader);
        updater.setValue("ia5", 1.5, 1.489);
        updater.setValue("ib5", 1.5, 1.4439);
        updater.setValue("ic5", 1.5, 1.562);
        updater.setValue("ic5", 0.5, 0.489);
        updater.setValue("ia5", 6, 5.23);
        updater.setValue("ic1", 0.7, 0.7456);
        //reader.printInfo();
        reader.save();
    }
}
