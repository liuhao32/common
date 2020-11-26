package com.example.demo.annotation;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Workbook;

import java.text.DecimalFormat;
import java.text.SimpleDateFormat;

/**
 * Copyright (C), 2020
 *
 * @author: liuhao
 * @date: 2020/11/22 20:45
 * @description:
 */

public class Handler {

    public static Object getCellValue(Cell cell) {  //Workbook wb,
        Object columnValue;
        if (cell == null) {
            return null;
        }
        DecimalFormat df = new DecimalFormat("0");
        DecimalFormat nf = new DecimalFormat("0.00");
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                columnValue = cell.getStringCellValue();
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if ("@".equals(cell.getCellStyle().getDataFormatString())) {
                    columnValue = df.format(cell.getNumericCellValue());
                } else if ("General".equals(cell.getCellStyle().getDataFormatString())) {
                    columnValue = nf.format(cell.getNumericCellValue());
                } else {
                    columnValue = HSSFDateUtil.getJavaDate(cell.getNumericCellValue()).toLocaleString();
                }
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                columnValue = cell.getBooleanCellValue();
                break;
            case Cell.CELL_TYPE_BLANK:
                columnValue = "";
                break;
//            case Cell.CELL_TYPE_FORMULA:
//                // 格式单元格
//                FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
//                evaluator.evaluateFormulaCell(cell);
//                CellValue cellValue = evaluator.evaluate(cell);
//                columnValue = cellValue.getNumberValue();
//                break;
            default:
                columnValue = cell.toString();
        }

        return columnValue;
    }
}
