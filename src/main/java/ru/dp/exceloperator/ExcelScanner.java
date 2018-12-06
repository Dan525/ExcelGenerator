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
    private Map<String, TypeCursor> targetMap;

    public ExcelScanner(String EXCEL_TEMPLATE) throws IOException {
        this.EXCEL_TEMPLATE = EXCEL_TEMPLATE;
        init(EXCEL_TEMPLATE);
    }

    private void init(String EXCEL_TEMPLATE) throws FileNotFoundException, IOException {
        FileInputStream is = new FileInputStream(new File(EXCEL_TEMPLATE));
        book = new XSSFWorkbook(is);
        targetMap = new HashMap<>();
    }

    public void iterateOnCells(int sheetNum) {
        XSSFSheet sheet = book.getSheetAt(0);

        Iterator<Row> rowIter = sheet.iterator();

        while (rowIter.hasNext()) {
            Row row = rowIter.next();
            Iterator<Cell> cellIter = row.cellIterator();

            while (cellIter.hasNext()) {
                Cell cell = cellIter.next();

                if (cell.getCellTypeEnum().equals(CellType.STRING)) {
                    String cellValue = cell.getStringCellValue();
                    findStartTypes(cellValue, cell);
                    findEndTypes(cellValue, cell);
                    if (cellValue.startsWith("target_")) {
                        CellAddress cellCoord = cell.getAddress();

                        if (cellValue.endsWith("_end")) {

                            String typeName = cellValue
                                    .replace("target_", "")
                                    .replace("_end", "");

                            if (targetMap.get(typeName) == null) {
                                throw new IndexOutOfBoundsException("end не может находиться раньше начала");
                            }
                            targetMap.get(typeName).setEndAddress(cellCoord);
                        } else {
                            String typeName = cellValue
                                    .replace("target_", "");
                            if (targetMap.get(typeName) == null) {
                                targetMap.put(typeName, new TypeCursor());
                            }
                            targetMap.get(typeName).setStartAddress(cellCoord);
                        }
                    }
                }

            }
        }
    }

    public void printInfo() {
        if (targetMap != null) {
            for (Map.Entry<String, TypeCursor> entry : targetMap.entrySet()) {
                System.out.println(entry.getKey());
                System.out.println("Число измерений: " + (entry.getValue().getEndAddress().getRow() + 1
                        - entry.getValue().getStartAddress().getRow()));
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
    public void findStartTypes(String str, Cell cell) {

        Pattern pattern = Pattern.compile("((nom)\\_([a-z0-9]+)\\_)[^end]?\\d+"); //nom_"какой-то тип"_не должно быть end, должна быть последовательность цифр
        Matcher matcher = pattern.matcher(str);
        if (matcher.find()) {
            System.out.println(matcher.group(3) + " " + cell.getAddress().getRow());
            cell.setCellValue(str.replaceAll(matcher.group(1), ""));
        }
    }

    public void findEndTypes(String str, Cell cell) {
        Pattern pattern = Pattern.compile("(nom)\\_([a-z0-9]+)\\_(end)\\_");
        Matcher matcher = pattern.matcher(str);
        if (matcher.find()) {
            System.out.println(matcher.group(2) + " " + cell.getAddress().getRow());
            cell.setCellValue(str.replaceAll(pattern.pattern(), ""));
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
