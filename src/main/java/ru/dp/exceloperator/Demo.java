/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ru.dp.exceloperator;

import java.io.IOException;

/**
 *
 * @author daniil_pozdeev
 */
public class Demo {
    public static void main(String[] args) throws IOException {
        ExcelScanner reader = new ExcelScanner("D:\\MyWorkspace\\protocol.xlsx");
        reader.iterateOnCells(0);
        reader.printInfo();
    }
}
