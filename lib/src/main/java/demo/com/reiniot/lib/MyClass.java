package demo.com.reiniot.lib;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class MyClass {
    private static List<Row> delta_two = new ArrayList<>();
    private static List<Row> delta_one = new ArrayList<>();
    private static List<Row> common = new ArrayList<>();
    private static int lastIndex = -1;

    private final static String ROOT_PATH = "D:\\data\\";
    public static void main(String[] args) throws Exception {

//        FileInputStream stream2 = getStream("D:\\data\\mobile.xlsx");
//        FileInputStream stream1 = getStream("D:\\data\\mobile_all.xlsx");
//        getDelataRow(stream1, 3, stream2, 0);

    }

    static private FileInputStream getStream(String path) {
        File file = new File(path);
        FileInputStream stream = null;
        try {
            stream = new FileInputStream(file);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        return stream;
    }


    static private void getDelataRow(FileInputStream stream, int dentifyIndex, FileInputStream stream2, int dentifyIndex2) {

        XSSFWorkbook workbook = null;
        XSSFWorkbook workbook2 = null;
        try {
            workbook = new XSSFWorkbook(stream);
            workbook2 = new XSSFWorkbook(stream2);
        } catch (IOException e) {
            e.printStackTrace();
        }
        XSSFSheet sheet = workbook.getSheetAt(0);
        XSSFSheet sheet2 = workbook2.getSheetAt(0);
        int rowsCount = sheet.getPhysicalNumberOfRows();
        int rowsCount2 = sheet2.getPhysicalNumberOfRows();
        delta_one.add(sheet.getRow(0));
        common.add(sheet.getRow(0));
        for (int r = 1; r < rowsCount; r++) {
            Row row = sheet.getRow(r);
            for (int i = 1; i < rowsCount2; i++) {
                Row row2 = sheet2.getRow(i);
                String id = getIdentifer(row, dentifyIndex).trim();
                String id2 = getIdentifer(row2, dentifyIndex2).trim();
                System.out.println("sheet1 row =" + r + ",sheet1 id =" + id + ",sheet2 id = " + id2 + "\n");
                if (id.equals(id2)) {

                    break;
                }

                if (i == rowsCount2 - 1) {
                    //添加sheet1中特有的行
                    delta_one.add(row);
                }

            }
        }

        for (int r2 = 1; r2 < rowsCount2; r2++) {
            Row row2 = sheet2.getRow(r2);
            for (int r1 = 1; r1 < rowsCount; r1++) {
                Row row1 = sheet.getRow(r1);
                String id = getIdentifer(row1, dentifyIndex).trim();
                String id2 = getIdentifer(row2, dentifyIndex2).trim();
                if (id.equals(id2)) {
                    common.add(row2);
                    break;
                }
                if (r1 == rowsCount - 1) {
                    //添加sheet1中特有的行
                    delta_two.add(row2);
                }

            }
        }

        saveNewExcel(ROOT_PATH+"delta2.xlsx",delta_two);
        saveNewExcel(ROOT_PATH+"delta1.xlsx",delta_one);
        saveNewExcel(ROOT_PATH+"common.xlsx", common);
    }

    private static void saveNewExcel(String path,List<Row> data){
        XSSFWorkbook wbCreat = new XSSFWorkbook();
        Sheet sheet_test = wbCreat.createSheet("test");

        for (int i = 0; i <data.size() ; i++) {
            Row t = sheet_test.createRow(sheet_test.getLastRowNum() +1);
            createCell(t, data.get(i));
        }

        saveExcel(wbCreat, path );
    }

    /**
     * 在已有的Excel文件中插入一行新的数据的入口方法
     */
    public static void insertRows(XSSFWorkbook wb, XSSFSheet sheet, Row data) {
        Row row = sheet.createRow((short) (lastIndex)); //在现有行号后追加数据
        createCell(row, data);
        lastIndex++;
    }

    /**
     * 找到需要插入的行数，并新建一个POI的row对象
     *
     * @param sheet
     * @param rowIndex
     * @return
     */
    private static XSSFRow createRow(XSSFSheet sheet, Integer rowIndex) {
        XSSFRow row = null;
        if (sheet.getRow(rowIndex) != null) {
            int lastRowNo = sheet.getLastRowNum();
            sheet.shiftRows(rowIndex, lastRowNo, 1);
        }
        row = sheet.createRow(rowIndex);
        return row;
    }

    /**
     * 创建要出入的行中单元格
     *
     * @param row
     * @return
     */
    private static void createCell(Row row, Row data) {
        for (int j = 0; j < data.getPhysicalNumberOfCells(); j++) {
            Cell cd = data.getCell(j);
            if (cd == null)
                continue;
            cd.setCellType(Cell.CELL_TYPE_STRING);
            String value = cd.getStringCellValue();
            //   System.out.println("value = " +value);
            row.createCell(j).setCellValue(value);
        }

    }

    static private String getIdentifer(Row row, int dentify) {
        String value = getCellAsString(row, dentify, null);
        return value;
    }

    static private void addBackCell(XSSFWorkbook wb, Cell cell) {
        CellStyle style = wb.createCellStyle();
        style.setFillForegroundColor((short) 2);
        style.setFillPattern((short) 1);
        cell.setCellStyle(style);
    }

    static private void saveExcel(XSSFWorkbook wb, String path) {
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

    static protected String getCellAsString(Row row, int c, FormulaEvaluator formulaEvaluator) {
        String value = "";
        try {
            Cell cell = row.getCell(c);
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


}
