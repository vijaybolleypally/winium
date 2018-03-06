
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import util.SheetUtill;

import java.io.IOException;

public class SheetsUtillTest {
    public static void main(String[] args) throws IOException, InvalidFormatException {

        new SheetUtill().compareExcel("C:\\Users\\E002961\\Desktop\\exs\\Expected.xlsx",
                "C:\\Users\\E002961\\Desktop\\exs\\Actual.xlsx", "C:\\Users\\E002961\\Desktop\\exs\\Result.xlsx");
    }
}