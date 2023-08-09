package com.example.excel_utis.controller;

import com.example.excel_utis.domain.Farmer;
import com.example.excel_utis.excel.ExportUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;


@RestController
public class FarmerController {

    @Resource
    private HttpServletResponse response;

    @GetMapping("/export")
    public void export() {
        // 模拟数据
        List<Farmer> farmerList = new ArrayList<>();
        Farmer farmer1 = new Farmer("宋剑寒","1","123456");
        Farmer farmer2 = new Farmer("unclesong","2","654321");
        farmerList.add(farmer1);
        farmerList.add(farmer2);
        // 表头数据中文数据集合
        List<String> headerList = new ArrayList<>();
        headerList.add("养殖户姓名");
        headerList.add("身份证类型");
        headerList.add("身份证号");

        // 调用把实体类集合转换成 excel 表
        ExportUtils.exportExcel("农户信息.xlsx", farmerList, Farmer.class, headerList, response);
    }

    @PostMapping("/import")
    public void importExcel(@RequestParam("excel") MultipartFile excel) throws IOException, InvalidFormatException {
        Workbook workbook = WorkbookFactory.create(excel.getInputStream());
        List<Farmer> farmerList = ExportUtils.importExcel(workbook, Farmer.class);
        System.out.println(farmerList);
    }
}
