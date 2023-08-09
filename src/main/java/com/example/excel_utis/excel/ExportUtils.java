package com.example.excel_utis.excel;


import org.apache.commons.codec.binary.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.CollectionUtils;


import javax.servlet.http.HttpServletResponse;
import java.lang.reflect.Field;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;

/**
 * &#064;Author     ：songjianhan
 * &#064;Date       : 2023/8/9 15:50
 * &#064;Description: 此类为 Excel 表格的导出类， 实体类对象集合 --> excel文件（.xlsx文件）
 */

public class ExportUtils {

    /**
     * 将传入的对象集合按行写入 Excel 文件中
     * eg: ExportUtils.exportExcel("农户信息.xlsx", farmerList, Farmer.class, response)
     * @param fileName 文件名, 不用带后缀
     * @param entityList 实体类对象集合
     * @param entityClass 实体类.class
     * @param headerListInChinese 中文表头集合
     * @param response HttpServletResponse
     */
    public static void exportExcel(String fileName, List<?> entityList, Class<?> entityClass, List<String> headerListInChinese, HttpServletResponse response) {
        try {
            // 创建 Excel 对象, 获取 sheet 对象来操作单表
            Workbook workBook = new XSSFWorkbook();
            Sheet sheet = workBook.createSheet();

            // 判空
            if (entityClass == null || CollectionUtils.isEmpty(entityList)) {
                throw new RuntimeException("class 对象或数据不能为空！");
            }

            // 通过反射获取实体类的属性名
            Row rowHeader = sheet.createRow(0);
            Field[] declaredFields = entityClass.getDeclaredFields();
            List<String> headerList = new ArrayList<>(); // 表头集合(英文)
            if (declaredFields.length == 0) {
                return;
            }
            //  遍历实体类的属性名集合
            for (int i = 0; i < declaredFields.length; i++) {
                Cell cell = rowHeader.createCell(i, CellType.STRING);
                String headerName = declaredFields[i].getName();
                // 强转一次 String
                String nameInChinese = headerListInChinese.get(i);
                cell.setCellValue(nameInChinese); // 填写中文表头
                headerList.add(i, headerName); // 记录表头的下标和表头的名字
            }

            for (int o = 0; o < entityList.size(); o++) {
                // 创建一行
                Row rowData = sheet.createRow(o + 1);
                for (int i = 0; i < headerList.size(); i++) {
                    // 填写这一行的数据
                    Cell cell = rowData.createCell(i);
                    Field nameField = entityClass.getDeclaredField(headerList.get(i)); // 通过反射使用属性名来写入数据
                    nameField.setAccessible(true);
                    String value = String.valueOf(nameField.get(entityList.get(o)));
                    cell.setCellValue(value);
                }
            }

            // excel 文件传回浏览器
            response.setContentType("application/vnd.ms-excel");
            String resultFileName = URLEncoder.encode(fileName, String.valueOf(StandardCharsets.UTF_8));
            response.setHeader("Content-disposition", "attachment;filename=" + resultFileName + ";" + "filename*=utf-8''" + resultFileName);
            workBook.write(response.getOutputStream());
            workBook.close();
            response.flushBuffer();

        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 从 Excel 中导入数据  Excel --> 实体类集合
     * 使用举例：前端使用 Post 表单上传文件，文件名必须为 excel
     *          后端controller代码应为：public void importExcel(@RequestParam("excel") MultipartFile excel)
     *          使用WorkbookFactory构造对象：Workbook workbook = WorkbookFactory.create(excel.getInputStream());
     *          调用：ExportUtils.importExcel(workbook, Farmer.class);
     * @param workbook 工作簿，请使用 WorkbookFactory 传入 inputStream 来获取
     * @param entity 实体类
     * @return 实体类集合
     */
    public static <T> List<T> importExcel(Workbook workbook, Class<?> entity){
        List<T> dataList = new ArrayList<>();
        try {
            Sheet sheet = workbook.getSheetAt(0);
            T o = null;
            for (Row row : sheet) {
                if(row != null){
                    o = (T) entity.getDeclaredConstructor().newInstance();
                    // 获取表头
                    Field[] declaredFields = entity.getDeclaredFields();
                    for (int i = 0; i < declaredFields.length; i++) {
                        // 通过反射来找到实体类和表中类的对应关系
                        String name = declaredFields[i].getName();
                        Field declaredField1 = o.getClass().getDeclaredField(name);
                        declaredField1.setAccessible(true);
                        Cell cell = row.getCell(i);
                        String type = declaredFields[i].getType().getName();
                        String value = String.valueOf(cell);
                        // 通过反射来获取对应类属性的类型
                        if(StringUtils.equals(type,"int") || StringUtils.equals(type,"Integer")){
                            declaredField1.set(o,Integer.parseInt(value));
                        } else if(StringUtils.equals(type,"java.lang.String") || StringUtils.equals(type,"char") || StringUtils.equals(type,"Character") ||
                                StringUtils.equals(type,"byte") || StringUtils.equals(type,"Byte")){
                            declaredField1.set(o,value);
                        } else if(StringUtils.equals(type,"boolean") || StringUtils.equals(type,"Boolean")){
                            declaredField1.set(o,Boolean.valueOf(value));
                        } else if(StringUtils.equals(type,"double") || StringUtils.equals(type,"Double")){
                            declaredField1.set(o,Double.valueOf(value));
                        } else if (StringUtils.equals(type,"long") || StringUtils.equals(type,"Long")) {
                            declaredField1.set(o,Long.valueOf(value));
                        } else if(StringUtils.equals(type,"short") || StringUtils.equals(type,"Short")){
                            declaredField1.set(o,Short.valueOf(value));
                        } else if(StringUtils.equals(type,"float") || StringUtils.equals(type,"Float")){
                            declaredField1.set(o,Float.valueOf(value));
                        }
                    }
                }
                dataList.add(o);
            }
            workbook.close();
            dataList.remove(0);
            return dataList;
        }catch (Exception e){
            e.printStackTrace();
        }
        return dataList;
    }


}
