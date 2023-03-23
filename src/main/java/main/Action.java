package main;

import util.ReadExcel;
import util.WriteExcel;

import java.util.LinkedHashMap;
//生成考号
public class Action {
    public static void main(String[] args) {
        //文档要求：按科目组升序排序，数字保存为文本。
        String filePath="E:\\隆回\\隆回一中\\考号生成\\考室安排.xlsx";
        //文档要求：按科目组升序，总分降序排序。
        String tofilePath="E:\\隆回\\隆回一中\\考号生成\\编排考号.xlsx";
        ReadExcel readExcel = new ReadExcel();
        LinkedHashMap<String, Integer> kaoShi = readExcel.getKaoShi(filePath, "");
        WriteExcel writeExcel = new WriteExcel();
        writeExcel.writeTestNum(tofilePath,kaoShi);

    }
}
