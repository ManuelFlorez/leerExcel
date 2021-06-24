package create;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class Excel {
    public static void main(String[] args) {
        createExcel();
        createExcelV2();
    }

    public static void createExcel() {
        Workbook book = new HSSFWorkbook();
        Sheet sheet = book.createSheet("Name Sheet");
        try {
            FileOutputStream fileOut = new FileOutputStream("data.xls");
            book.write(fileOut);
            fileOut.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void createExcelV2() {
        Workbook book = new XSSFWorkbook();
        Sheet sheet = book.createSheet("Name Sheet");
        try {
            FileOutputStream fileOut = new FileOutputStream("data.xlsx");
            book.write(fileOut);
            fileOut.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
