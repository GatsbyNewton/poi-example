package com.wzm.example;

import com.wzm.tool.ExcelUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ListTemplate {
    private static final Logger LOGGER = LoggerFactory.getLogger(ListTemplate.class);

    public void createHeader(Workbook workbook, Sheet sheet, List<String> headers){
        Row headerRow = sheet.createRow(0);     //第一个sheet的第一行为标题
        sheet.createFreezePane(0, 1, 0, 1);     //冻结第一行

        for (int i = 0, size = headers.size(); i < size; i++){
            Cell cell = headerRow.createCell(i);
            cell.setCellStyle(ExcelUtils.setHeaderCellStyle(workbook, sheet));
            cell.setCellValue(headers.get(i));
        }
    }


    public void createThreeLevelDependentDropDownList(Workbook workbook, Sheet sheet, List<String> l1DropDownList,
                                            Map<String, List<String>> dependentDropDownListMap, int firstRow,
                                            int lastRow, int firstCol, int lastCol){
        ExcelUtils.createHiddenSheet(workbook, l1DropDownList, dependentDropDownListMap, "hidden", "test");
        ExcelUtils.setThreeLevelDropDownListDependency(sheet, l1DropDownList, firstRow, lastRow, firstCol, lastCol);
    }

    private void extractExcelTemplate(String path)throws Exception{
        try {
            List<String> headers = Arrays.asList("设备名称", "设类型", "数量", "存放库房", "存放货架", "存放层", "存放列");
            List<String> devices = Arrays.asList("设备1", "设备2", "设备3");
            Workbook workbook = ExcelUtils.createWorkBook(path);
            Sheet sheet = workbook.createSheet("test");
            createHeader(workbook, sheet, headers);
            ExcelUtils.createDropDownList(sheet, devices, 1, 6, 0, 6);

            try (OutputStream fileOut = new FileOutputStream(path)) {
                fileOut.flush();
                workbook.write(fileOut);
            }
            System.out.println("导出成功!");
        }
        catch (IOException ioe){
            throw ioe;
        }
    }

    private void extractThreeLevelDropDownListExcelTemplate(String path)throws Exception{
        try {
            Workbook workbook = ExcelUtils.createWorkBook(path);
            Sheet sheet = workbook.createSheet("test");
            List<String> headers = Arrays.asList("姓名", "省", "市", "区");

            List<String> provices = Arrays.asList("江苏省", "安徽省");
            List<String> cityJiangSu = Arrays.asList("南京市", "苏州市", "盐城市");
            List<String> cityAnHui = Arrays.asList("合肥市", "安庆市");
            List<String> countyNanjing = Arrays.asList("六合县", "江宁县");
            List<String> countySuzhou = Arrays.asList("姑苏区", "园区");
            List<String> countyYancheng = Arrays.asList("响水县", "射阳县");
            List<String> countyLiuhe = Arrays.asList("瑶海区", "庐阳区");
            List<String> countyAnQing = Arrays.asList("迎江区", "大观区");
            //将有子区域的父区域放到一个数组中
            Map<String, List<String>> areaMap = new HashMap<>();
            areaMap.put("江苏省", cityJiangSu);
            areaMap.put("安徽省", cityAnHui);
            areaMap.put("南京市", countyNanjing);
            areaMap.put("苏州市", countySuzhou);
            areaMap.put("盐城市", countyYancheng);
            areaMap.put("合肥市", countyYancheng);
            areaMap.put("合肥市", countyLiuhe);
            areaMap.put("安庆市", countyAnQing);
            createHeader(workbook, sheet, headers);
            createThreeLevelDependentDropDownList(workbook, sheet, provices, areaMap, 1, 5, 1, 1);

            try (OutputStream fileOut = new FileOutputStream(path)) {
                fileOut.flush();
                workbook.write(fileOut);
            }
            System.out.println("导出成功!");
        }
        catch (IOException ioe){
            throw ioe;
        }
    }

    public static void main(String... args)throws Exception{
        ListTemplate template = new ListTemplate();
        String path = "/Users/gatsbynewton/Documents/codes/java/poi-example/data/";
//        template.extractExcelTemplate(path + "test.xls");
        template.extractThreeLevelDropDownListExcelTemplate(path + "test2.xlsx");
    }
}
