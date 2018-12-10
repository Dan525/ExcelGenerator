/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ru.dp.exceloperator;

/**
 *
 * @author daniil_pozdeev
 */
public class ExcelUpdater {
    
    private ExcelScanner scanner;

    public ExcelUpdater(ExcelScanner scanner) {
        this.scanner = scanner;
    }

    public void setValue(String type, double nomValue, double measValue) {
        int y = scanner.getMeasTypeMap().get(type).getRow();
        int xCell = scanner.getMeasTypeMap().get(type).getColumn();
        
        int yCell = scanner.getNominalMap().get(y).getNominalValuesMap().get(nomValue);
        
        scanner.getBook().getSheetAt(0).getRow(yCell).getCell(xCell).setCellValue(measValue);
    }
}
