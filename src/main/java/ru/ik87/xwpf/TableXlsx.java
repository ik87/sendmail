package ru.ik87.xwpf;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class TableXlsx {
    final String inputFile;

    public TableXlsx(String tableXlsxFile) {
        this.inputFile = tableXlsxFile;
    }

    public List<Map<String, String>> readXlsx() throws IOException, InvalidFormatException {

        List<Map<String, String>> tables = new ArrayList<>();
        List<String> variables = new ArrayList<>();

        try (OPCPackage pkg = OPCPackage.open(new File(inputFile))) {
            XSSFWorkbook wb = new XSSFWorkbook(pkg);
            FormulaEvaluator objFormulaEvaluator = new XSSFFormulaEvaluator(wb);
            Sheet sheet1 = wb.getSheetAt(0);
            for (Row row : sheet1) {
                Map<String, String> table = new LinkedHashMap<>();
                for (Cell cell : row) {
                    if (row.getRowNum() == 0) {
                        variables.add(cell.getRichStringCellValue().getString());
                    } else {
                        if (cell.getCellType() == CellType.STRING) {

                            String value = cell.getRichStringCellValue().getString();
                            String key = variables.get(cell.getColumnIndex());
                            table.put(key, value);

                        }

                        if (cell.getCellType() == CellType.FORMULA) {
                            DataFormatter df = new DataFormatter();
                            String key = variables.get(cell.getColumnIndex());
                            objFormulaEvaluator.evaluate(cell);
                            String value = df.formatCellValue(cell, objFormulaEvaluator);
                            //System.out.println(value);
                            table.put(key, value);
                        }
                    }
                }
                if (row.getRowNum() != 0) {
                    tables.add(table);
                }
            }
        }
        return tables;
    }

    public void writeXlsx(int indexRow, String flag, String desc) throws IOException {
        FileInputStream fis = new FileInputStream(new File(inputFile));
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        Sheet sheet1 = wb.getSheetAt(0);
        Map<String, String> table = new LinkedHashMap<>();
        for (Cell cell : sheet1.getRow(0)) {
                if (Objects.equals(cell.getRichStringCellValue().getString(), flag)) {
                    int indexColumn = cell.getColumnIndex();
                    Row r = sheet1.getRow(indexRow);
                    Cell c = r.getCell(indexColumn);
                    if(c == null) {
                        c = r.createCell(indexColumn);
                    }
                    c.setCellValue(desc);
                }
        }
        fis.close();
        FileOutputStream fos = new FileOutputStream(new File(inputFile));
        wb.write(fos);
        fos.close();
    }
}
