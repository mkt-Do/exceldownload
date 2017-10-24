package org.demo.Util;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.web.servlet.view.document.AbstractXlsxView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class ExcelBuilderUtil extends AbstractXlsxView {

    /**
     * ExcelBuilder
     * How to Use
     * 1. headerMap(key is number(from 0), value is List<String>)
     * 2. contentMap(key is number(from 0), value is List<Object>)
     * 3. sheetList(List<String>)
     * 4. add to Map above elements(key is String(headers, contents, sheets), value is above elements)
     * 5. use method 
     *    ModelAndView mav = new ModelAndView(new ExcelBuilderUtil(), map);
     */
    @Override
    protected void buildExcelDocument(Map<String, Object> model, Workbook workbook, HttpServletRequest request, HttpServletResponse response) throws Exception {

        // headerList
        Map<Integer, List<String>> headerMap = (Map<Integer, List<String>>) model.get("headers");

        // contentMap
        Map<Integer, Object> contentMap = (Map<Integer, Object>) model.get("contents");

        // sheetNameList
        List<String> sheetList = (List<String>) model.get("sheets");
        
        for(int num = 0; num < sheetList.size(); num++) {
            
            // create sheet
            Sheet sheet = workbook.createSheet(sheetList.get(num));
            
            List<String> headerList = headerMap.get(num);
            
            // fetch contentList from contentMap
            List<Object> contentList = (List<Object>) contentMap.get(num);

            // add header
            Row row = sheet.createRow(0);
            for(int i = 0; i < headerList.size(); i++) {
                row.createCell(i).setCellValue(headerList.get(i));
            }

            // get class
            if(contentList.isEmpty()) {
                break;
            }
            Class clazz = contentList.get(0).getClass();

            // get fields
            Field[] fields = isStaticField(clazz.getDeclaredFields());

            for(int i = 0; i < contentList.size(); i++) {
                row = sheet.createRow(i + 1);

                Object content = contentList.get(i);
                for(int j = 0; j < fields.length; j++) {

                    // create getter
                    String fieldName = fields[j].getName();
                    Method method = clazz.getMethod("get" + toUpperCase(fieldName));

                    // invoke getter method
                    String result = String.valueOf(method.invoke(content));
                    if(isLong(result)) {
                        // result is long number
                        row.createCell(j).setCellValue(Long.parseLong(result));
                    } else if(isDouble(result)) {
                        // result is double number
                        row.createCell(j).setCellValue(Double.parseDouble(result));
                    } else {
                        // others
                        row.createCell(j).setCellValue(result);
                    }
                }
            }
        }

    }

    /**
     * drop static field
     * @param fields
     * @return
     */
    private Field[] isStaticField(Field[] fields) {

        List<Field> fieldList = new ArrayList<>();
        for(Field field : fields) {
            if(!Modifier.isStatic(field.getModifiers())) {
                fieldList.add(field);
            }
        }
        return fieldList.toArray(new Field[fieldList.size()]);
    }
    
    /**
     * fetch getter(lombok)
     * @param str
     * @return
     */
    static private String toUpperCase(String str) {
        return Character.toTitleCase(str.charAt(0)) + str.substring(1);
    }

    private boolean isLong(String value) {
        try {
            long num = Long.parseLong(value);

            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }

    private boolean isDouble(String value) {
        try {
            Double num = Double.parseDouble(value);

            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }
}