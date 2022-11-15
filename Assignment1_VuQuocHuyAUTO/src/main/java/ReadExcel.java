import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {

    public static void main(String[] args) {

        String path = "D:\\Telegram Desktop\\test_data\\Assignment1_VuQuocHuyAUTO\\src\\main\\resources\\testdata.xlsx";

        try {

            //Create an object of FileInputStream class to read excel file
            FileInputStream fis = new FileInputStream(path);

            //Create object of XSSFWorkbook class
            XSSFWorkbook wb = new XSSFWorkbook(fis);

            //Read excel sheet by sheet name
            XSSFSheet sheet1 = wb.getSheet("Sheet1");
            DataFormatter dataFormatter = new DataFormatter();
            for (int i = 4; i <= sheet1.getLastRowNum(); i++) {
                String cellValue = dataFormatter.formatCellValue(sheet1.getRow(i).getCell(3));
                //Get data from specified cell
                List<String> list = new ArrayList<String>();
                list.add(cellValue);
                System.out.println(list.size());
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

    }

}