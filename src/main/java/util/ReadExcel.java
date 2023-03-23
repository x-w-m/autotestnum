package util;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.util.LinkedHashMap;


public class ReadExcel {
    public Workbook getWorkbook(String filePath, String password) {
        Workbook workbook;
        File file = new File(filePath);
        try {
            workbook = WorkbookFactory.create(file, password, true);
            return workbook;
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    public LinkedHashMap<String, Integer> getKaoShi(String filePath, String password) {
        Workbook workbook = getWorkbook(filePath, password);
        LinkedHashMap<String, Integer> linkedHashMap = new LinkedHashMap<>();
        Sheet sheet = workbook.getSheetAt(0);
        //最大行号下标
        int lastRowNum = sheet.getLastRowNum();
        for (int i = 0; i <= lastRowNum; i++) {
            System.out.println(lastRowNum);
            Row row = sheet.getRow(i);
            //考室号
            Cell cell = row.getCell(0);
            String kaoshi = cell.getStringCellValue();
            System.out.println(kaoshi);
            //座位号
            cell = row.getCell(1);
            Integer zuowei = (int)cell.getNumericCellValue();
            System.out.println(zuowei);
            linkedHashMap.put(kaoshi, zuowei);
        }

        return linkedHashMap;
    }

}

