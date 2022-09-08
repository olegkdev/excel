package excel;

import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;

public class ReadExcelTest {
    String excelFileAbsolutePath = "C:\\Projects\\excel\\file.xlsx";
    int sheetIndex = 0;
    Integer[] cellsIndexes = new Integer[]{1, 2, 3};
    int rowsNumber = 4;
    int rowStart = 1;
    int rowEnd = 2;

    @Test
    void testRead() throws IOException {
        List<List<?>> results = ReadExcel.read(this.excelFileAbsolutePath, this.sheetIndex, this.cellsIndexes);

        assertEquals(results.size(), this.rowsNumber);

        for (List<?> result : results) {
            assertEquals(result.size(), this.cellsIndexes.length);
        }
    }

    @Test
    void testReadRange() throws IOException {
        List<List<?>> results = ReadExcel.readRange(this.excelFileAbsolutePath, this.sheetIndex, this.rowStart,
                this.rowEnd, this.cellsIndexes);

        assertEquals(results.size(), this.rowEnd - this.rowStart + 1);

        for (List<?> result : results) {
            assertEquals(result.size(), this.cellsIndexes.length);
        }
    }
}
