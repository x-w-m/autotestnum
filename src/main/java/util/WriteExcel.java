package util;

import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.LinkedHashMap;

public class WriteExcel {
    public Workbook getWorkbook(String filePath, String password) {
        Workbook workbook;
        File file = new File(filePath);
        try {
            workbook = WorkbookFactory.create(file, password, false);
            return workbook;
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    public boolean writeTestNum(String filePath, LinkedHashMap<String, Integer> linkedHashMap) {
        File file = new File(filePath);
        InputStream is;
        OutputStream os;
        Workbook workbook = null;
        try {
            is = new FileInputStream(file);
            workbook = WorkbookFactory.create(is);
            Sheet sheet = workbook.getSheetAt(0);
            //使用引用变量记录数据条数
            int[] arr = {2};
            linkedHashMap.forEach((k, v) -> {
                for (int i = 1; i <= v; i++) {

                    Row row = sheet.getRow(arr[0]);
                    Cell cell = row.createCell(8);
                    cell.setCellValue(k);
                    cell = row.createCell(9);
                    String zuoweihao;
                    if (i < 10) {
                        zuoweihao = "0" + i;
                    } else zuoweihao = "" + i;
                    cell.setCellValue(zuoweihao);
                    arr[0]++;
                    System.out.println(k + "" + zuoweihao);
                }


            });
            os = new FileOutputStream(file);
            workbook.write(os);
            os.flush();
            os.close();
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return false;
    }

    public boolean createRow(String filePath, int sheetnum) {
        File file = new File(filePath);
        InputStream is;
        OutputStream os;
        Workbook workbook = null;
        try {
            is = new FileInputStream(file);
            workbook = WorkbookFactory.create(is);
            Sheet sheet = workbook.getSheetAt(sheetnum);
            //表格最后一行在变化，每次循环需要重新获取。
            for (int i = 10; true; i++) {

                Row row = sheet.getRow(i);
                Cell cell = row.getCell(6);
                String upstr = cell.getStringCellValue();
                Row row1 = sheet.getRow(i + 1);
                //循环结束条件：不存在下一行或下一行座位号为空
                if (row1 == null || row1.getCell(6) == null || row1.getCell(6).getStringCellValue().equals("")) {
                    break;
                }
                Cell cell1 = row1.getCell(6);
                String downstr = cell1.getStringCellValue();
                System.out.println(upstr + ":" + downstr);
                if (Integer.parseInt(upstr) > Integer.parseInt(downstr)) {
                    //偶数考室添加2行，奇数考室添加3行。
                    if (Integer.parseInt(upstr) % 2 == 0) {
                        //将后面的数据下移两行
                        sheet.shiftRows(i + 1, sheet.getLastRowNum(), 2);
                        sheet.createRow(++i).createCell(4).setCellValue("0");
                        sheet.createRow(++i).createCell(4).setCellValue("0");
                    } else {
                        //将后面的数据下移三行
                        sheet.shiftRows(i + 1, sheet.getLastRowNum(), 3);
                        sheet.createRow(++i).createCell(4).setCellValue("0");
                        sheet.createRow(++i).createCell(4).setCellValue("0");
                        sheet.createRow(++i).createCell(4).setCellValue("0");
                    }
                }
            }
            os = new FileOutputStream(file);
            workbook.write(os);
            os.flush();
            os.close();
            workbook.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return false;
    }
}
