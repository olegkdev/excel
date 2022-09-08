package excel;

import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Locale;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

public class WriteExcelTest {
    static List<Data> objects;
    static List<String> fieldsNames;
    String fileName = "file";
    String sheetName = "objects";

    @BeforeAll
    static void init() throws ParseException {
        WriteExcelTest.fieldsNames = new ArrayList<String>() {{
            add("name");
            add("date");
            add("number");
            add("bool");
        }};

        WriteExcelTest.objects = new ArrayList<Data>() {{
            SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd", Locale.ENGLISH);
            String dateInString = "1032-10-01";
            Date date = formatter.parse(dateInString);
            add(new Data("name1", date, 10, true));

            dateInString = "2032-12-02";
            date = formatter.parse(dateInString);
            add(new Data("name2", date, 20, true));

            dateInString = "3032-06-29";
            date = formatter.parse(dateInString);
            add(new Data("name3", date, 30, false));

            dateInString = "4032-08-11";
            date = formatter.parse(dateInString);
            add(new Data("name4", date, 40, false));
        }};
    }

    @Test
    void testWrite() throws IOException, NoSuchFieldException, IllegalAccessException {
        String fileLocation = WriteExcel.write(this.fileName, this.sheetName, WriteExcelTest.objects,
                WriteExcelTest.fieldsNames);

        File file = new File(fileLocation);
        assertTrue(file.exists());
        assertFalse(file.isDirectory());
    }

    static class Data {
        String name;
        Date date;
        int number;
        boolean bool;

        Data(String name, Date date, int number, boolean bool) {
            this.name = name;
            this.date = date;
            this.number = number;
            this.bool = bool;
        }
    }
}
