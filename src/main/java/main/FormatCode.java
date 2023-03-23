package main;

import util.WriteExcel;

public class FormatCode {
    public static void main(String[] args) {
        String filePath="E:\\隆回\\隆回一中\\考号生成\\条形码格式化.xlsx";
        WriteExcel writeExcel = new WriteExcel();
        //i表示工作表下标，共五个工作表
        for(int i=0;i<5-4;i++){
            writeExcel.createRow(filePath,i);
        }

    }
}
