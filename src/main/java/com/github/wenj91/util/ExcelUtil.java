package com.github.wenj91.util;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * Created by wenj91 on 2017/3/17.
 */
public class ExcelUtil {
    private static final Logger log = Logger.getLogger(ExcelUtil.class);

    public static Workbook exportXLSX(Map<String, Object> mapper, List<Map<String, Object>> items){
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Sheet 1");
        try {
            int currentRow = 0;
            currentRow = fillItemHeaderRowByFields(sheet, mapper, currentRow);
            for(Map<String, Object> item : items){
                Row row = sheet.createRow(currentRow++);
                fillMapperPropertiesWithMap(item, row, mapper.keySet());
            }
        } finally {
            log.info("Create XLSX Success!!!");
        }

        return wb;
    }

    public static Workbook exportXLSX(Map<String, Object> mapper, JSONArray items){
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Sheet 1");
        try {
            int currentRow = 0;
            currentRow = fillItemHeaderRowByFields(sheet, mapper, currentRow);
            for(Object itemObj : items){
                JSONObject itemJson = (JSONObject) itemObj;
                Row row = sheet.createRow(currentRow++);
                fillMapperPropertiesWithJSON(itemJson, row, mapper.keySet());
            }
        } finally {
            log.info("Create XLSX Success!!!");
        }

        return wb;
    }


    private static void fillMapperPropertiesWithJSON(JSONObject jo, Row row, Set<String> customFields) {
        int currentCell = 0;
        for (String field : customFields){
            row.createCell(currentCell++).setCellValue(jo.getString(field));
        }
    }

    private static void fillMapperPropertiesWithMap(Map<String, Object> item, Row row, Set<String> customFields) {
        int currentCell = 0;
        for (String field : customFields){
            row.createCell(currentCell++).setCellValue(item.get(field)==null?"":String.valueOf(item.get(field)));
        }
    }

    private static int fillItemHeaderRowByFields(Sheet sheet, Map<String, Object> mapper, int currentRow) {
        Set<String> fields = mapper.keySet();
        Set<String> mapperFields = new LinkedHashSet<>();
        for(String field:fields){
            mapperFields.add((String) mapper.get(field));
        }
        Row headerRow = sheet.createRow(currentRow++);
        int currentCell = 0;
        for(String field : mapperFields){
            headerRow.createCell(currentCell++).setCellValue(field);
        }
        return currentRow;
    }
}
