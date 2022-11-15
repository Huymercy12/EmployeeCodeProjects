import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

public class ExcelUtils {


    public static List<ExcelDataRowModel> readFileExcel(String pathFile) throws IOException, InvalidFormatException {
        FileInputStream file = new FileInputStream(new File(pathFile));

        XSSFWorkbook workbook = new XSSFWorkbook(file);

        XSSFSheet sheet = workbook.getSheetAt(0);
        Iterator<Row> rowIterator = sheet.iterator();
        List<ExcelDataRowModel> lstRs = new ArrayList<>();
        DataFormatter fmt = new DataFormatter();
        int i = 3;
        while (rowIterator.hasNext()) {
            ExcelDataRowModel rowExcel = new ExcelDataRowModel();
            rowExcel.setRowIndex(i);

            Row row = rowIterator.next();
            Iterator<Cell> cellIterator = row.cellIterator();

            List<Object> rowData = new ArrayList<>();

            while (cellIterator.hasNext()) {
                Object objData = null;

                Cell cell = cellIterator.next();
                // Đổi thành getCellType() nếu sử dụng POI 4.x
                switch (cell.getCellType()) {
                    case STRING:
                        objData = fmt.formatCellValue(cell);
                        break;

                }
                rowData.add(objData);
            }
            rowExcel.setDataRows(rowData);
            lstRs.add(rowExcel);
            i++;
        }

        return lstRs;
    }

    public static void main(String[] args) throws IOException, InvalidFormatException {
        List<ExcelDataRowModel> lstData = ExcelUtils.readFileExcel("D:\\Telegram Desktop\\test_data\\Assignment1_VuQuocHuyAUTO\\src\\main\\resources\\testdata.xlsx");
        System.out.println(lstData.size());
        int b = 0;
        for (int i = 4; i < lstData.size(); i++) {
            String c = String.valueOf(lstData.get(i).getDataRows().get(3));
            List<String> namelist = new ArrayList<>();
            namelist.add(String.valueOf(c));
            List<Object> namelist2 = new ArrayList<>();
            for (Object element : namelist) {
                if (!namelist2.contains(element)) {
                    namelist2.add(element);
                }
            }
            System.out.println(namelist.get(0));
        }
    }
}


//hien thi du lieu ra
//            for (ExcelDataRowModel dataRows : lstData) {
//                List<Object> cellDatas = dataRows.getDataRows();
//                if (cellDatas.get(3) == "Phi Thi Thom") {
//                    System.out.println(cellDatas.get(3));
//                }
//            }



