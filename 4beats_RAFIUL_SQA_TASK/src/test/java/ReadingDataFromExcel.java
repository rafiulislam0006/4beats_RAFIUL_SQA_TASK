import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class ReadingDataFromExcel {

    public static void main(String[] args) throws IOException {

        FileInputStream file = new FileInputStream("C://Users//rafi//Desktop//test.xlsx");

        XSSFWorkbook workbook = new XSSFWorkbook(file);

        XSSFSheet sheet = workbook.getSheet("Saturday");

        int rowcount = sheet.getLastRowNum(); // return the row count

        int colcount = sheet.getRow(1).getLastCellNum(); // return the cell count

        for (int i = 1; i < rowcount; i++){
            XSSFRow currentrow = sheet.getRow(i); // focused on current row

            for (int j = 1; j < colcount; j++){
                String value =  currentrow.getCell(j).toString(); // read the value from a cell
                System.out.print((" " + value));
            }
            System.out.println();
        }

    }
}
