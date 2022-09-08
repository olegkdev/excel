package excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ReadExcel {
    public static List<List<?>> read(String excelFileAbsolutePath, int sheetIndex, Integer[] cellsIndexes)
            throws IOException {
        FileInputStream file = new FileInputStream(new File(excelFileAbsolutePath));
        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(sheetIndex);

        List<List<?>> results = new ArrayList<>();

        for (Row row : sheet) {
            List<Object> result = new ArrayList<>();

            for (int cellIndex : cellsIndexes) {
                Cell cell = row.getCell(cellIndex);
                CellType cellType = cell.getCellType();

                switch (cellType) {
                    case NUMERIC:
                        result.add(cell.getNumericCellValue());
                        break;

                    case STRING:
                        result.add(cell.getStringCellValue());
                        break;

                    case BOOLEAN:
                        result.add(cell.getBooleanCellValue());
                        break;

                    default:
                        result.add(null);
                }
            } // for (int cellIndex : cellsIndexes)

            results.add(result);
        } // for (Row row : sheet)

        return results;
    }

    public static List<List<?>> readRange(String excelFileAbsolutePath, int sheetIndex, int rowStart, int rowEnd,
                                          Integer[] cellsIndexes)
            throws IOException {
        FileInputStream file = new FileInputStream(new File(excelFileAbsolutePath));
        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(sheetIndex);

        List<List<?>> results = new ArrayList<>();

        for (int rowIndex = rowStart; rowIndex <= rowEnd; rowIndex++) {
            List<Object> result = new ArrayList<>();

            for (int cellIndex : cellsIndexes) {
                Row row = sheet.getRow(rowIndex);
                Cell cell = row.getCell(cellIndex);
                CellType cellType = cell.getCellType();

                switch (cellType) {
                    case NUMERIC:
                        result.add(cell.getNumericCellValue());
                        break;

                    case STRING:
                        result.add(cell.getStringCellValue());
                        break;

                    case BOOLEAN:
                        result.add(cell.getBooleanCellValue());
                        break;

                    default:
                        result.add(null);
                }
            } // for (int cellIndex : cellsIndexes)

            results.add(result);
        } // for (int rowIndex = rowStart; rowIndex <= rowEnd; rowIndex++)

        return results;
    }
}
