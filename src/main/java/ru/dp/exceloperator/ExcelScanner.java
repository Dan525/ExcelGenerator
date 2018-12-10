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
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.NoSuchElementException;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author daniil_pozdeev
 */
public class ExcelScanner {

    private final String EXCEL_TEMPLATE;
    private final String EXCEL_COMPLETED;

    private XSSFWorkbook book;
    private Map<String, CellAddress> measTypeMap;
    private Map<Integer, NominalValues> nominalMap;

    public ExcelScanner(String templatePath, String completedPath) throws IOException {
        EXCEL_TEMPLATE = templatePath;
        EXCEL_COMPLETED = completedPath;
        init(EXCEL_TEMPLATE);
    }

    private void init(String templatePath) throws FileNotFoundException, IOException {
        FileInputStream is = new FileInputStream(new File(templatePath));
        book = new XSSFWorkbook(is);
        measTypeMap = new HashMap<>();
        nominalMap = new HashMap<>();
    }

    public void scanSheet(int sheetNum) {
        XSSFSheet sheet = book.getSheetAt(0);
        Iterator<Row> rowIter = sheet.iterator();

        while (rowIter.hasNext()) {
            Row row = rowIter.next();
            Iterator<Cell> cellIter = row.cellIterator();
            while (cellIter.hasNext()) {
                Cell cell = cellIter.next();
                checkNom(cell);
            }
        }
    }

    /*
        http://www.quizful.net/post/Java-RegExp
        Круглые скобки - группа
        Квадратные скобки - допустимые символы
        ^ - отрицание
        ? - предшествующий символ может быть, а может не быть
        \\d - цифры
        + - последовательность из предшествующего символа
     */
    private void checkNom(Cell cell) {
        switch (cell.getCellTypeEnum()) {
            case STRING: {
                processString(cell);
                break;
            }                
            case FORMULA: {
                processFormula(cell);
            }                
        }
    }

    private void processString(Cell cell) {
        String str = cell.getStringCellValue();
        Pattern pattern = Pattern.compile("([Nn]om\\_)(-?[0-9]+,?[0-9]*)");
        Matcher matcher = pattern.matcher(str);
        if (matcher.find()) {
            double value = Double.valueOf(matcher.group(2).replace(",", "."));
            cell.setCellType(CellType.NUMERIC);
            cell.setCellValue(value);
            int y = cell.getAddress().getRow();
            if (nominalMap.get(y) == null) {
                nominalMap.put(y, new NominalValues());
            }
            findTypesInRow(cell);
            fillNomValues(cell);
        }
    }

    private void processFormula(Cell cell) {
        String formula = cell.getCellFormula();
        Pattern pattern = Pattern.compile("\"[Nn]om_\" *& *");
        Matcher matcher = pattern.matcher(formula);
        if (matcher.find()) {
            formula = formula.replaceAll(matcher.group(0), "");
            cell.setCellFormula(formula);
            int y = cell.getAddress().getRow();
            if (nominalMap.get(y) == null) {
                nominalMap.put(y, new NominalValues());
            }
            findTypesInRow(cell);
            fillNomValues(cell);
        }
    }

    private void fillNomValues(Cell cell) {
        int y = cell.getAddress().getRow();
        int x = cell.getAddress().getColumn();
        NominalValues values = nominalMap.get(y);
        XSSFSheet sheet = book.getSheetAt(0);
        boolean end = false;

        for (int i = 0;; i++) {
            Cell iterableCell = sheet.getRow(y + i).getCell(x);
            if (end || iterableCell == null) {
                break;
            }
            switch (iterableCell.getCellTypeEnum()) {
                case NUMERIC: {
                    double value = iterableCell.getNumericCellValue();
                    values.getNominalValuesMap().put(value, iterableCell.getAddress().getRow());
                    break;
                }
                case FORMULA: {
                    String str1 = iterableCell.getRichStringCellValue().toString();
                    double doub = iterableCell.getNumericCellValue();
                    if (iterableCell.getCachedFormulaResultTypeEnum().equals(CellType.STRING)) {
                        String str = iterableCell.getStringCellValue();
                        double value = Double.valueOf(iterableCell.getStringCellValue().replace(",", "."));
                        values.getNominalValuesMap().put(value, iterableCell.getAddress().getRow());
                        break;
                    }
                }
                case BLANK: {
                    end = true;
                }
            }
        }
    }

    private void checkMeasurementType(Cell cell) {
        if (cell.getCellTypeEnum().equals(CellType.STRING)) {
            String str = cell.getStringCellValue();
            Pattern pattern = Pattern.compile("([Mm]eas\\_)([a-z0-9]+)");
            Matcher matcher = pattern.matcher(str);
            if (matcher.find()) {
                String measType = matcher.group(2);
                cell.setCellType(CellType.BLANK);
                int y = cell.getAddress().getRow();
                int x = cell.getAddress().getColumn();
                measTypeMap.put(measType, new CellAddress(y, x));
            }
        }

    }

    private void findTypesInRow(Cell cell) {
        int y = cell.getAddress().getRow();
        XSSFSheet sheet = book.getSheetAt(0);
        Iterator<Cell> cellIter = sheet.getRow(y).cellIterator();
        while (cellIter.hasNext()) {
            Cell iterableCell = cellIter.next();
            checkMeasurementType(iterableCell);
        }
    }
    
    public void setValue(String type, double nomValue, double measValue) {
        if (measTypeMap.get(type) == null) {
            throw new NoSuchElementException("No such type: " + type);
        }
        int y = measTypeMap.get(type).getRow();
        int xCell = measTypeMap.get(type).getColumn();
        
        Map<Double, Integer> nominalValuesMap = nominalMap.get(y).getNominalValuesMap();
        if (nominalValuesMap.get(nomValue) == null) {
            throw new NoSuchElementException("No such nominal value for this type: " + type);
        }
        int yCell = nominalMap.get(y).getNominalValuesMap().get(nomValue);
        
        book.getSheetAt(0).getRow(yCell).getCell(xCell).setCellValue(measValue);
    }

    public void save() {
        File file = new File(EXCEL_COMPLETED);
        if (!file.exists()) {
            try {
                file.createNewFile();
            } catch (IOException ex) {
                Logger.getLogger(ExcelScanner.class.getName()).log(Level.SEVERE, "Can't create file", ex);
            }
        }
        try {
            FileOutputStream out = new FileOutputStream(file);
            book.write(out);
            out.close();
        } catch (FileNotFoundException ex) {
            Logger.getLogger(ExcelScanner.class.getName()).log(Level.SEVERE, "File not found", ex);
        } catch (IOException e) {
            Logger.getLogger(ExcelScanner.class.getName()).log(Level.SEVERE, "Can't write book", e);
        }
    }

    public Map<String, CellAddress> getMeasTypeMap() {
        return measTypeMap;
    }

    public Map<Integer, NominalValues> getNominalMap() {
        return nominalMap;
    }

    public XSSFWorkbook getBook() {
        return book;
    }
}
