/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ru.dp.exceloperator;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.ThreadLocalRandom;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author daniil_pozdeev
 */
public class ExcelUpdater {

    public static void main(String[] args) throws FileNotFoundException, IOException {
        File file = new File("D:\\MyWorkspace\\protocol.xlsx");
        FileInputStream is = new FileInputStream(file);
        XSSFWorkbook book = new XSSFWorkbook(is);
        XSSFSheet sheet = book.getSheetAt(0);
        for (int rowNum = 30; rowNum < 39; rowNum++) {
            for (int cellNum = 4; cellNum < 7; cellNum++) {
                XSSFCell cell = sheet.getRow(rowNum).getCell(cellNum);
                double randomValue = ThreadLocalRandom.current().nextDouble(50, 210);
                cell.setCellValue(randomValue);
            }
        }
        is.close();

        FileOutputStream out = new FileOutputStream(file);
        book.write(out);
        out.close();
    }

}
