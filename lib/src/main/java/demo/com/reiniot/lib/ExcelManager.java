package demo.com.reiniot.lib;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * 2019/4/30.
 */
public class ExcelManager {
    private static ExcelManager singleInstance = null;
    private final static String ROOT_PATH = "D:\\data\\";
    private String small;
    private String big;
    private FileInputStream streamSmall;
    private FileInputStream streamBig;
    private int smallDentifyId = -1;
    private int bigDentifyId = -1;

    private static List<Row> delta_big = new ArrayList<>();
    private static List<Row> delta_small = new ArrayList<>();
    private static List<Row> common = new ArrayList<>();

    public static synchronized ExcelManager getSingleInstance() {
        if (singleInstance == null) {
            singleInstance = new ExcelManager();
        }
        return singleInstance;
    }


    public ExcelManager setSmallFilePath(String small) {
        if (small != null && small.length() > 0) {
            this.small = small;
            try {
                streamSmall = new FileInputStream(small);
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            }
            return this;
        }
        return null;
    }

    public ExcelManager setBigFilePath(String big) {
        if (big != null && big.length() > 0) {
            this.big = big;
            try {
                streamBig = new FileInputStream(big);
            } catch (FileNotFoundException e) {
                e.printStackTrace();
                return null;
            }
            return this;
        }
        return null;

    }

    public ExcelManager setBigDentifyId(int bigDentifyId) {
        this.bigDentifyId = bigDentifyId;
        return this;
    }

    public ExcelManager setSmallDentifyId(int smallDentifyId) {
        this.smallDentifyId = smallDentifyId;
        return this;
    }

    public void excute() {
        getDelataRow(streamSmall, smallDentifyId, streamBig, bigDentifyId);

    }

    private void getDelataRow(FileInputStream small, int sd, FileInputStream big, int bd) {

        XSSFWorkbook smallWorkbook = null;
        XSSFWorkbook bigWorkbook = null;
        try {
            smallWorkbook = new XSSFWorkbook(small);
            bigWorkbook = new XSSFWorkbook(big);
        } catch (IOException e) {
            e.printStackTrace();
            return;
        }
        XSSFSheet smallSheet = smallWorkbook.getSheetAt(0);
        XSSFSheet bigSheet = bigWorkbook.getSheetAt(0);
        int smallRowCount = smallSheet.getPhysicalNumberOfRows();
        int bigRowCount = bigSheet.getPhysicalNumberOfRows();
        //添加excel header到新的excel文件中
        delta_small.add(bigSheet.getRow(0));
        common.add(bigSheet.getRow(0));
        delta_big.add(bigSheet.getRow(0));

        for (int r = 1; r < smallRowCount; r++) {
            Row row = smallSheet.getRow(r);
            for (int i = 1; i < bigRowCount; i++) {
                Row row2 = bigSheet.getRow(i);
                String id = getIdentifer(row, sd).trim();
                String id2 = getIdentifer(row2, bd).trim();
                System.out.println("sheet1 row =" + r + ",sheet1 id =" + id + ",sheet2 id = " + id2 + "\n");
                if (id.equals(id2)) {

                    break;
                }

                if (i == bigRowCount - 1) {
                    //添加sheet1中特有的行
                    delta_small.add(row);
                }

            }
        }

        for (int r2 = 1; r2 < bigRowCount; r2++) {
            Row row2 = bigSheet.getRow(r2);
            for (int r1 = 1; r1 < smallRowCount; r1++) {
                Row row1 = smallSheet.getRow(r1);
                String id = getIdentifer(row1, sd).trim();
                String id2 = getIdentifer(row2, bd).trim();
                if (id.equals(id2)) {
                    common.add(row2);
                    break;
                }
                if (r1 == smallRowCount - 1) {
                    //添加sheet1中特有的行
                    delta_big.add(row2);
                }

            }
        }

        saveNewExcel(ROOT_PATH + "delta_big.xlsx", delta_big, "delta_big_sheet");
        saveNewExcel(ROOT_PATH + "delta_small.xlsx", delta_small, "delta_small_sheet");
        saveNewExcel(ROOT_PATH + "common.xlsx", common, "common_sheet");
    }


    private String getIdentifer(Row row, int dentify) {
        String value = getCellAsString(row, dentify);
        return value;
    }


    private String getCellAsString(Row row, int index) {
        String value = "";
        try {
            Cell cell = row.getCell(index);
            cell.setCellType(Cell.CELL_TYPE_STRING);
            if (null != cell) {
                switch (cell.getCellType()) {                     // 判断excel单元格内容的格式，并对其进行转换，以便插入数据库
                    case 0:
                        value = String.valueOf((int) cell.getNumericCellValue());
                        break;
                    case Cell.CELL_TYPE_STRING:
                        value = cell.getStringCellValue();
                        break;
                    case 2:
                        value = cell.getNumericCellValue() + "";
                        // cellValue = String.valueOf(cell.getDateCellValue());
                        break;
                    case 3:
                        value = "";
                        break;
                    case 4:
                        value = String.valueOf(cell.getBooleanCellValue());
                        break;
                    case 5:
                        value = String.valueOf(cell.getErrorCellValue());
                        break;
                }
            } else {
                value = "";
            }
        } catch (NullPointerException e) {
        }
        return value;
    }

    private void saveNewExcel(String path, List<Row> data, String sheetName) {
        XSSFWorkbook wbCreat = new XSSFWorkbook();
        if (sheetName == null)
            sheetName = "test";
        Sheet sheet_test = wbCreat.createSheet(sheetName);

        for (int i = 0; i < data.size(); i++) {
            Row t = sheet_test.createRow(sheet_test.getLastRowNum() + 1);
            createCell(t, data.get(i));
        }

        saveExcel(wbCreat, path);
    }

    private void createCell(Row row, Row data) {
        for (int j = 0; j < data.getPhysicalNumberOfCells(); j++) {
            Cell cd = data.getCell(j);
            if (cd == null)
                continue;
            cd.setCellType(Cell.CELL_TYPE_STRING);
            String value = cd.getStringCellValue();
            row.createCell(j).setCellValue(value);
        }
    }

    private void saveExcel(XSSFWorkbook wb, String path) {
        FileOutputStream fileOut;
        try {
            fileOut = new FileOutputStream(path);
            wb.write(fileOut);
            fileOut.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}
