package com.wzm.tool;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Map;
import java.util.Set;

public class ExcelUtils {
    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelUtils.class);

    /**
     * 创建一个专门用来存放地区信息的隐藏sheet页, 因此也不能在现实页之前创建，否则无法隐藏。
     * @param workbook
     * @param l1DropDownList 父下拉列表
     * @param dependentDropDownListMap 级联列映射
     * @param hiddenSheetName 隐藏sheet页名称
     * @param l1Value 隐藏sheet页中父下拉列表的title
     */
    public static void createHiddenSheet(Workbook workbook, List<String> l1DropDownList,
                                         Map<String, List<String>> dependentDropDownListMap,
                                         String hiddenSheetName, String l1Value){
        Sheet hiddenSheet = workbook.createSheet(hiddenSheetName);
        /* 这一行作用是将此sheet隐藏，功能未完成时注释此行,可以查看隐藏sheet中信息是否正确 */
        //book.setSheetHidden(book.getSheetIndex(hideSheet), true);

        int rowId = 0;
        Row l1Row = hiddenSheet.createRow(rowId++);
        l1Row.createCell(0).setCellValue(l1Value);
        for (int i = 0, size = l1DropDownList.size(); i < size; i++){
            Cell cell = l1Row.createCell(i + 1);
            cell.setCellValue(l1DropDownList.get(i));
        }

        // 将具体的数据写入到每一行中，行开头为父级区域，后面是子区域。
        Set<String> fatherSet = dependentDropDownListMap.keySet();
        int fatherIdx = 0;
        for (String father : fatherSet){
            List<String> sons = dependentDropDownListMap.get(father);
            Row row = hiddenSheet.createRow(rowId++);
            row.createCell(0).setCellValue(father);

            for (int i = 0, size = sons.size(); i < size; i++){
                Cell cell = row.createCell(i + 1);
                cell.setCellValue(sons.get(i));
            }

            /* 添加名称管理器 */
            String range = getRange(1, rowId, sons.size());
            Name name = workbook.createName();
            // key不可重复
            name.setNameName(father);
            String formula = hiddenSheetName + "!" + range;
            name.setRefersToFormula(formula);
        }
    }

    /**
     * 设置三级级联列
     * @param sheet 主sheet页
     * @param l1DropDownList 父下拉列表
     * @param firstRow 父下拉列表生效起始行，从0开始计数
     * @param lastRow 父下拉列表生效终止行，从0开始计数
     * @param firstCol 父下拉列表生效起始列，从0开始计数
     * @param lastCol 父下拉列表生效终止列，从0开始计数
     */
    public static void setThreeLevelDropDownListDependency(Sheet sheet, List<String> l1DropDownList, int firstRow,
                                                           int lastRow, int firstCol, int lastCol){
        DataValidationHelper dvHelper = sheet.getDataValidationHelper();
        DataValidationConstraint l1ListConstraint = dvHelper.createExplicitListConstraint(l1DropDownList.toArray(new String[]{}));
        CellRangeAddressList l1ListRangeAddressList = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
        DataValidation l1ListDataValidation = dvHelper.createValidation(l1ListConstraint, l1ListRangeAddressList);
        l1ListDataValidation.createErrorBox("error", "请选择正确的省份");
        l1ListDataValidation.setEmptyCellAllowed(false);
        if(l1ListDataValidation instanceof XSSFDataValidation) {
            l1ListDataValidation.setSuppressDropDownArrow(true);
            l1ListDataValidation.setShowErrorBox(true);
        }
        else {
            l1ListDataValidation.setSuppressDropDownArrow(false);
        }
        sheet.addValidationData(l1ListDataValidation);

        /* 设置子列表有效性 */
        for (int i = firstRow + 1; i < lastRow + 2; i++) {
            setDataValidation(String.valueOf((char) ('A' + firstCol)), sheet, i, firstCol + 1);
            setDataValidation(String.valueOf((char) ('A' + firstCol + 1)), sheet, i, firstCol + 2);
        }
    }

    /**
     * 设置级联有效性
     * @param offset 级联列偏移
     * @param sheet
     * @param rowNum 级联列生效行
     * @param colNum 级联列生效列
     */
    private static void setDataValidation(String offset, Sheet sheet, int rowNum, int colNum) {
        DataValidationHelper dvHelper = sheet.getDataValidationHelper();
        DataValidation dataValidation = getDataValidationByFormula(
                "INDIRECT($" + offset + (rowNum) + ")", rowNum, colNum, dvHelper);
        sheet.addValidationData(dataValidation);
    }

    /**
     * 设置级联
     * @param formulaString 映射函数
     * @param naturalRowIndex 级联列生效行
     * @param naturalColumnIndex 级联列生效列
     * @param dvHelper
     * @return
     */
    private static DataValidation getDataValidationByFormula(String formulaString, int naturalRowIndex,
                                                             int naturalColumnIndex, DataValidationHelper dvHelper) {
        /**
         * 加载下拉列表内容,
         * 举例：若formulaString = "INDIRECT($A$2)" 表示规则数据会从名称管理器中获取key与单元格 A2 值相同的数据，
         * 如果A2是江苏省，那么此处就是江苏省下的市信息。
         */
        DataValidationConstraint dvConstraint = dvHelper.createFormulaListConstraint(formulaString);
        int firstRow = naturalRowIndex - 1;
        int lastRow = naturalRowIndex - 1;
        int firstCol = naturalColumnIndex;
        int lastCol = naturalColumnIndex;
        CellRangeAddressList regions = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
        DataValidation dataValidation = dvHelper.createValidation(dvConstraint, regions);
        dataValidation.createPromptBox("下拉选择提示", "请使用下拉方式选择合适的值！");
        dataValidation.setEmptyCellAllowed(false);

        if (dataValidation instanceof XSSFDataValidation) {
            dataValidation.setSuppressDropDownArrow(true);
            dataValidation.setShowErrorBox(true);
        } else {
            dataValidation.setSuppressDropDownArrow(false);
        }
        // 设置输入信息提示信息
        return dataValidation;
    }

    /**
     * 计算formula
     * @param offset 偏移量，如果给0，表示从A列开始，1，就是从B列
     * @param rowId 第几行
     * @param colCount 一共多少列
     * @return 如果给入参 1,1,10. 表示从B1-K1。最终返回 $B$1:$K$1
     */
    public static String getRange(int offset, int rowId, int colCount) {
        char start = (char) ('A' + offset);
        if (colCount <= 25) {
            char end = (char) (start + colCount - 1);
            return "$" + start + "$" + rowId + ":$" + end + "$" + rowId;
        } else {
            char endPrefix = 'A';
            char endSuffix = 'A';
            if ((colCount - 25) / 26 == 0 || colCount == 51) {// 26-51之间，包括边界（仅两次字母表计算）
                if ((colCount - 25) % 26 == 0) {// 边界值
                    endSuffix = (char) ('A' + 25);
                } else {
                    endSuffix = (char) ('A' + (colCount - 25) % 26 - 1);
                }
            } else {// 51以上
                if ((colCount - 25) % 26 == 0) {
                    endSuffix = (char) ('A' + 25);
                    endPrefix = (char) (endPrefix + (colCount - 25) / 26 - 1);
                } else {
                    endSuffix = (char) ('A' + (colCount - 25) % 26 - 1);
                    endPrefix = (char) (endPrefix + (colCount - 25) / 26);
                }
            }
            return "$" + start + "$" + rowId + ":$" + endPrefix + endSuffix + "$" + rowId;
        }
    }

    public static boolean validateExcel(String fileName){
        return fileName.endsWith(".xls") || fileName.endsWith(".xlsx");
    }

    public static Workbook createWorkBook(InputStream in){
        Workbook workbook = null;
        try {
            workbook = WorkbookFactory.create(in);
        }
        catch (IOException ioe){
            LOGGER.error("", ioe);
        }
        catch (InvalidFormatException ife){
            LOGGER.error("", ife);
        }

        return workbook;
    }

    public static Workbook createWorkBook(String fileName){
        if (fileName.endsWith(".xls")) {
            return new HSSFWorkbook();
        }
        else {
            return new XSSFWorkbook();
        }
    }

    /**
     * 设置标题样式
     * @param workbook
     * @param sheet
     */
    public static CellStyle setHeaderCellStyle(Workbook workbook, Sheet sheet){
        CellStyle cellStyle = setCellStyle(workbook.createCellStyle());

        Font font = setFont(workbook.createFont());
        font.setBold(true);
        cellStyle.setFont(font);

        return cellStyle;
    }

    /**
     * 设置数据样式
     * @param workbook
     * @param sheet
     */
    public static CellStyle setDataCellStyle(Workbook workbook, Sheet sheet){
        CellStyle cellStyle = setCellStyle(workbook.createCellStyle());

        Font font = setFont(workbook.createFont());
        cellStyle.setFont(font);

        return cellStyle;
    }

    /**
     * 设置错误样式
     * @param workbook
     * @param sheet
     */
    public static CellStyle setErrorCellStyle(Workbook workbook, Sheet sheet){
        CellStyle cellStyle = setCellStyle(workbook.createCellStyle());

        Font font = setFont(workbook.createFont());
        font.setBold(true);
        font.setColor(Font.COLOR_RED);
        cellStyle.setFont(font);

        return cellStyle;
    }

    private static CellStyle setCellStyle(CellStyle cellStyle){
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setWrapText(true);

        return cellStyle;
    }

    private static Font setFont(Font font){
        font.setFontName("宋体");

        return font;
    }

    public static void main(String... args){
    }
}
