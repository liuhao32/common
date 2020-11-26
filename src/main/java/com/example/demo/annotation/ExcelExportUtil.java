package com.example.demo.annotation;

import lombok.Getter;
import lombok.Setter;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;

import java.io.FileOutputStream;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Collectors;

/**
 * Copyright (C), 2020
 *
 * @author: liuhao
 * @date: 2020/11/21 19:25
 * @description:
 */

public class ExcelExportUtil {


    /**
     * @param output
     * @param pojoList
     * @throws Exception
     */
    public static void exportExcel(FileOutputStream output, List<?> pojoList, List<List<Integer>> failPos,
                                   List<List<String>> failCause) throws Exception {

        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.createSheet("错误数据导出表");
        CellStyle style = wb.createCellStyle();
        style.setFillForegroundColor(HSSFColor.RED.index);
        style.setFillBackgroundColor(HSSFColor.RED.index);
        style.setFillPattern(HSSFCellStyle.BIG_SPOTS);
        HSSFRow row = sheet.createRow(0);
        Class<?> clazz = pojoList.get(0).getClass();

        Field[] declaredFields = clazz.getDeclaredFields();
        List<Excel> annotations = new ArrayList<>();
        for (Field field : declaredFields) {
            Excel excel = field.getAnnotation(Excel.class);
            if (excel != null) {
                annotations.add(excel);
            }
        }

        List<Excel> sortedExcel = annotations.stream().sorted(Comparator.comparing(Excel::column))
                .collect(Collectors.toList());

        List<String> tableRowName = sortedExcel.stream().map(Excel::name).collect(Collectors.toList());

        for (int i = 0; i < tableRowName.size(); i++) {
            row.createCell(i).setCellValue(tableRowName.get(i));
        }
        row.createCell(tableRowName.size()).setCellValue("原因");

        for (int i = 0; i < pojoList.size(); i++) {
            row = sheet.createRow(i + 1);
            Object pojo = pojoList.get(i);
            Class<?> pojoClazz = pojo.getClass();
            Field[] fs = pojoClazz.getDeclaredFields();

            List<Inter> inters = new ArrayList<>();
            for (Field f : fs) {
                Excel excel = f.getAnnotation(Excel.class);
                if (excel != null) {
                    f.setAccessible(true);
                    Object val = f.get(pojo);
                    int column = excel.column();
                    inters.add(new Inter(column, val));
                }
            }
            List<Inter> sortedInter = inters.stream().sorted(Comparator.comparing(Inter::getColumn))
                    .collect(Collectors.toList());

            List<Object> tableRowValue = sortedInter.stream().map(Inter::getVal).collect(Collectors.toList());

            List<Integer> pos = failPos.get(i);
            List<String> cause = failCause.get(i);

            // 创建单元格并填充数据
            for (int j = 0; j < tableRowValue.size(); j++) {
                HSSFCell cellJ = row.createCell(j);
                cellJ.setCellValue(String.valueOf(tableRowValue.get(j)));
                if (pos.contains(j)) {
                    cellJ.setCellStyle(style);
                }
            }
            row.createCell(tableRowValue.size()).setCellValue(String.join(",", cause));
        }
        //列宽自适应
        for (int i = 0; i <= tableRowName.size(); i++) {
            sheet.autoSizeColumn(i);
        }
        wb.write(output);
    }

    @Setter
    @Getter
    static class Inter {
        /**
         * 序列
         */
        int column;

        /**
         * 值
         */
        Object val;

        Inter(int column, Object val) {
            this.column = column;
            this.val = val;
        }
    }
}