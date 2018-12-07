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

    private XSSFWorkbook book;
    private Map<String, CellAddress> measTypeMap;
    private Map<Integer, NominalValues> nominalMap;

    public ExcelScanner(String EXCEL_TEMPLATE) throws IOException {
        this.EXCEL_TEMPLATE = EXCEL_TEMPLATE;
        init(EXCEL_TEMPLATE);
    }

    private void init(String EXCEL_TEMPLATE) throws FileNotFoundException, IOException {
        FileInputStream is = new FileInputStream(new File(EXCEL_TEMPLATE));
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

//    public void printInfo() {
//        if (targetMap != null) {
//            for (Map.Entry<String, NominalValues> entry : targetMap.entrySet()) {
//                System.out.println(entry.getKey());
//                System.out.println("Число измерений: " + (entry.getValue().getEndAddress().getRow() + 1
//                        - entry.getValue().getStartAddress().getRow()));
//            }
//        }
//    }

    /*
        http://www.quizful.net/post/Java-RegExp
        Круглые скобки - группа
        Квадратные скобки - допустимые символы
        ^ - отрицание
        ? - предшествующий символ может быть, а может не быть
        \\d - цифры
        + - последовательность из предшествующего символа
    */
//    public void findStartTypes(String str, Cell cell) {
//        Pattern pattern = Pattern.compile("((nom)\\_([a-z0-9]+)\\_)[^end]?\\d+"); //nom_"какой-то тип"_не должно быть end, должна быть последовательность цифр
//        Matcher matcher = pattern.matcher(str);
//        if (matcher.find()) {
//            System.out.println(matcher.group(3) + " " + cell.getAddress().getRow());
//            cell.setCellValue(str.replaceAll(matcher.group(1), ""));
//        }
//    }
//
//    public void findEndTypes(String str, Cell cell) {
//        Pattern pattern = Pattern.compile("(nom)\\_([a-z0-9]+)\\_(end)\\_");
//        Matcher matcher = pattern.matcher(str);
//        if (matcher.find()) {
//            System.out.println(matcher.group(2) + " " + cell.getAddress().getRow());
//            cell.setCellValue(str.replaceAll(pattern.pattern(), ""));
//        }
//    }
    
    private void checkNom(Cell cell) {
        String str = cell.getStringCellValue();
        Pattern pattern = Pattern.compile("(nom\\_)\\d+");
        Matcher matcher = pattern.matcher(str);
        if (matcher.find()) {
            cell.setCellValue(str.replaceAll(matcher.group(1), ""));
            int y = cell.getAddress().getRow();
            if (nominalMap.get(y) == null) {
                nominalMap.put(y, new NominalValues());
            }
            fillNomValues(cell);
            findTypesInRow(cell);
        }
    }
    
    private boolean checkNomEnd(Cell cell) {
        String str = cell.getStringCellValue();
        Pattern pattern = Pattern.compile("(nom\\_end\\_)\\d+");
        Matcher matcher = pattern.matcher(str);
        if (matcher.find()) {
            cell.setCellValue(str.replaceAll(matcher.group(1), ""));
            cell.setCellType(CellType.NUMERIC);
            return true;
        }
        return false;
    }
    
    private void fillNomValues(Cell cell) {
        int y = cell.getAddress().getRow();
        int x = cell.getAddress().getColumn();
        NominalValues values = nominalMap.get(y);
        XSSFSheet sheet = book.getSheetAt(0);
        
        for (int i = 0; ; i++) {
            Cell iterableCell = sheet.getRow(y + i).getCell(x);
            boolean nomEnd = checkNomEnd(iterableCell);
            if (iterableCell.getCellType().equals(CellType.NUMERIC)) {
                double value = iterableCell.getNumericCellValue();
                values.getNominalMap().put(value, iterableCell.getAddress().getRow());
            } else {
                Logger.getLogger(ExcelScanner.class.getName()).log(Level.WARNING, "Cell type is not numeric");
            }
            if (nomEnd) {
                break;
            }
        }
    }
    
    private void checkMeasurementType(Cell cell) {
        String str = cell.getStringCellValue();
        Pattern pattern = Pattern.compile("(meas\\_)([a-z0-9]+)"); //nom_"какой-то тип"_не должно быть end, должна быть последовательность цифр
        Matcher matcher = pattern.matcher(str);
        if (matcher.find()) {
            String measType = matcher.group(2);
            cell.setCellValue(str.replaceAll(matcher.group(0), ""));
            measTypeMap.put(measType, cell.getAddress());
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

    public void save() {
        File file = new File("D:\\MyWorkspace\\protocol_completed.xlsx");
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
}
