package excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Locale;

public class WriteExcel {
    public static String write(String fileName, String sheetName, List<?> objects, List<String> fieldsNames)
            throws IOException, NoSuchFieldException, IllegalAccessException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet(sheetName);

        for (int i = 0; i < objects.size(); i++) {
            Row row = sheet.createRow(i);

            for (int j = 0; j < fieldsNames.size(); j++) {
                Cell cell = row.createCell(j);
                Field field = (objects.get(i)).getClass().getDeclaredField(fieldsNames.get(j));
                field.setAccessible(true);
                Object value = field.get(objects.get(i));

                if (value instanceof Date) {
                    SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd", Locale.ENGLISH);
                    cell.setCellValue(formatter.format(value));
                } else if (value instanceof Integer) cell.setCellValue((int) value);
                else if (value instanceof Double) cell.setCellValue((double) value);
                else if (value instanceof String) cell.setCellValue((String) value);
                else if (value instanceof Boolean) cell.setCellValue((boolean) value);
            }
        }

        File currDir = new File(".");
        String path = currDir.getAbsolutePath();
        String fileLocation = path.substring(0, path.length() - 1) + fileName + ".xlsx";

        FileOutputStream outputStream = new FileOutputStream(fileLocation);
        workbook.write(outputStream);
        workbook.close();

        return fileLocation;
    }
}
