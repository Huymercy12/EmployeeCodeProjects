import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

public class EmployeeData {

    private static final String FILE_EXCEL = "D:\\Telegram Desktop\\test_data\\Assignment1_VuQuocHuyAUTO\\src\\main\\resources\\testdata.xlsx";
    private static final String DATA_WRITE = "D:\\Telegram Desktop\\test_data\\Assignment1_VuQuocHuyAUTO\\src\\main\\resources\\data.txt";
    private static int countID = 0;

    public static void main(String[] args) {
        Map<String, String> idAndName = new HashMap<String, String>();
        Map<String, Long> idAndTime = new HashMap<String, Long>();
        Map<String, Float> idAndDay = new HashMap<String, Float>();
        demNhanVien(idAndName, idAndTime, idAndDay);
        demNgay(idAndName, idAndTime, idAndDay);
        writeFile(idAndName, idAndTime, idAndDay);
    }

    public static void demNhanVien(Map<String, String> names, Map<String, Long> times, Map<String, Float> days) {
        try (FileInputStream fis = new FileInputStream(FILE_EXCEL)) {
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet mySheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = mySheet.iterator();
            while (rowIterator.hasNext()) {
                Row currRow = rowIterator.next();
                Cell id = currRow.getCell(2);
                Cell name = currRow.getCell(3);
                String currId = "", currName = "";
                if (name != null && id != null) {
                    if (id.getCellType() == CellType.STRING && name.getCellType() == CellType.STRING) {
                        currId = id.getStringCellValue();
                        currName = name.getStringCellValue();
                        if (currId.length() < 1 || currId.equalsIgnoreCase("ID")) {
                            countID++;
                        }
                    }
                    names.put(currId, currName);
                    times.put(currId, 0L);
                    days.put(currId, 0F);
                }
            }
            System.out.println("Successfully employee!!!");
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static boolean checkTime(Date dateCellValue, Date dateCellValue1, String checkInOut) throws ParseException {
        boolean check = true;

        try {
            if (checkInOut == "IN") {
                // CHECK IN
                if (dateCellValue.getTime() <= dateCellValue1.getTime()) {
                    check = true;
                } else {
                    check = false;
                }
            } else if (checkInOut == "OUT") {
                // CHECK OUT
                if (dateCellValue.getTime() >= dateCellValue1.getTime()) {
                    check = true;
                } else {
                    check = false;
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return check;
    }


    public static void demNgay(Map<String, String> names, Map<String, Long> times, Map<String, Float> days) {
        try (FileInputStream fis = new FileInputStream(FILE_EXCEL)) {
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet mySheet = workbook.getSheetAt(0);
            Iterator<Row> it = mySheet.iterator();
            float countDay = 1;
            while (it.hasNext()) {
                Row currRow = it.next();
                // lay id nhan vien
                Cell id = currRow.getCell(2);
                // Check in and check out
                Cell timeIn = currRow.getCell(4);
                Cell timeOut = currRow.getCell(5);
                // Time check in and check out chuẩn theo ca (đang lấy theo khung 8h - 17h)
                Cell timeInCheck = currRow.getCell(6);
                Cell timeOutCheck = currRow.getCell(7);
                String currId = "";
                long time = 0;
                if (timeIn != null && timeOut != null && timeInCheck != null && timeOutCheck != null && timeIn.getCellType() == CellType.NUMERIC && timeOut.getCellType() == CellType.NUMERIC && id.getCellType() == CellType.STRING) {
                    currId = id.getStringCellValue();
                    // CHECK TIME ĐI MUỘN VỀ SỚM
                    if (!checkTime(timeIn.getDateCellValue(), timeInCheck.getDateCellValue(), "IN")) {
                        time += compareTime(timeIn.getDateCellValue(), timeInCheck.getDateCellValue());
                    }

                    if (!checkTime(timeOut.getDateCellValue(), timeOutCheck.getDateCellValue(), "OUT") && (checkDays(timeIn.getDateCellValue(), timeOut.getDateCellValue()) == 1F)) {
                        time += compareTime(timeOut.getDateCellValue(), timeOutCheck.getDateCellValue());
                    }

                    // GÁN LẠI GIÁ TRỊ CHO THỜI GIAN ĐI MUỘN VÀ NGÀY CÔNG SAU MỖI NGÀY ĐI LÀM (NẾU CÓ)
                    long timeTemp = times.get(currId);
                    timeTemp += time;

                    times.put(currId, timeTemp);

                    float dayTemp = days.get(currId);
                    dayTemp += checkDays(timeIn.getDateCellValue(), timeOut.getDateCellValue());

                    days.put(currId, dayTemp);
                }
            }
            workbook.close();
            System.out.println("Successfully count times and days.");
        } catch (IOException | ParseException e) {
            e.printStackTrace();
        }
    }

    private static long compareTime(Date dateCellValue, Date dateCellValue1) throws ParseException {
        SimpleDateFormat simpleDateFormat
                = new SimpleDateFormat("HH:mm:ss");

        // Calculating the difference in milliseconds
        long differenceInMilliSeconds
                = Math.abs(dateCellValue1.getTime() - dateCellValue.getTime());

        // Calculating the difference in Hours
        long differenceInHours
                = (differenceInMilliSeconds / (60 * 60 * 1000))
                % 24;

        // Calculating the difference in Minutes
        long differenceInMinutes
                = (differenceInMilliSeconds / (60 * 1000)) % 60;

        // Calculating the difference in Seconds
        long differenceInSeconds
                = (differenceInMilliSeconds / 1000) % 60;

        long sogiaydilammuon = ((differenceInHours * 60 * 60) + (differenceInMinutes * 60) + (differenceInSeconds));


        return sogiaydilammuon;
    }

    private static float checkDays(Date dateCellValue, Date dateCellValue1) throws ParseException {
        SimpleDateFormat simpleDateFormat
                = new SimpleDateFormat("HH:mm:ss");

        // Calculating the difference in milliseconds
        long differenceInMilliSeconds
                = Math.abs(dateCellValue1.getTime() - dateCellValue.getTime());

        // Calculating the difference in Hours
        long differenceInHours
                = (differenceInMilliSeconds / (60 * 60 * 1000))
                % 24;

        // Calculating the difference in Minutes
        long differenceInMinutes
                = (differenceInMilliSeconds / (60 * 1000)) % 60;

        // Calculating the difference in Seconds
        long differenceInSeconds
                = (differenceInMilliSeconds / 1000) % 60;

        // Printing the answer
        float check = differenceInHours + (float) (differenceInMinutes / 60) + (float) (differenceInSeconds / 3600);

        // Trường hợp làm buổi sáng sau đó đầu giờ chiều mới ra về -> nửa công
        if (check < 6) return 0.5F;
        else return 1.0F;
    }

    public static void writeFile(Map<String, String> names, Map<String, Long> times, Map<String, Float> days) {
        try {
            FileWriter myWriter = new FileWriter(DATA_WRITE);

            myWriter.write("Sum Employee is: " + (names.size() - 1 - countID) + "\n");
            myWriter.write("ID : NAME : DAYS : TIME ĐI MUỘN TRONG THÁNG (ĐÃ TRỪ 1H/1 THÁNG)" + "\n");
            for (Map.Entry<String, String> m : names.entrySet()) {
                if (m.getKey().length() < 1 || m.getKey().equalsIgnoreCase("ID")) {
                    continue;
                }
                // CHUYỂN TIME ĐI MUỘN THÀNH GIỜ -> MỖI THÁNG CÓ 60 PHÚT ĐI MUỘN NÊN TRỪ ĐI 1H
                float countTime = ((float) times.get(m.getKey()) / 3600) - 1F;
                myWriter.write(m.getKey() + " : " + m.getValue() + " : " + days.get(m.getKey()) + " : " + (countTime > 0 ? countTime : 0) + "\n");
            }
            myWriter.close();
            System.out.println("Ghi thành công vào file txt.");
        } catch (IOException e) {
            System.out.println("An error occurred.");
            e.printStackTrace();
        }
    }
}