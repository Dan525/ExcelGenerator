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
        String templatePath = "D:\\MyWorkspace\\protocol.xlsx";
        String completedPath = "D:\\MyWorkspace\\protocol_completed.xlsx";
        ExcelScanner reader = new ExcelScanner(templatePath, completedPath);
        reader.scanSheet(0);
        reader.setValue("ia5", 0.05, 1.489);
        reader.setValue("ib5", 0.05, 1.4439);
        reader.setValue("ic5", 0.05, 1.562);
        reader.setValue("ia1", 0.01, 0.454);
        reader.setValue("ib1", 0.01, 8.46);
        reader.setValue("ic1", 0.01, 97.412);
        reader.setValue("ic5", 0.5, 0.489);
        reader.setValue("ia5", 6, 5.23);
        reader.setValue("ic1", 0.7, 0.7456);
        reader.save();
    }
}
