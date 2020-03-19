package com.van589.poi.demo.controller;

import com.van589.poi.demo.entity.User;
import com.van589.poi.demo.utils.ExportPOIUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

@Controller
public class MainController {

    @Value("${file.path}")
    private String path;

    @GetMapping(value = {"", "index"})
    public String main() {
        return "main";
    }

    /**
     * 导出excel
     *
     * @param response
     * @return
     */
    @ResponseBody
    @GetMapping("export")
    public String exportExcel(HttpServletResponse response) throws IOException {
        String fileName = "创建的User";


        List<User> users = new ArrayList<User>();

        users.add(new User("1","张三","123"));
        users.add(new User("2","李四","123"));
        users.add(new User("3","王五","123"));

        // 列名
        String columnNames[] = {"ID", "名称", "密码"};
        // map中的key,即实体类中的字段名
        String keys[] = {"id", "name", "password"};
        ExportPOIUtils.start_download(response, fileName, users, columnNames, keys);
        return "OK";
    }

    /**
     * 实现文件上传
     */
    @RequestMapping("fileUpload")
    @ResponseBody
    public String fileUpload(@RequestParam("fileName") MultipartFile file) throws IOException {
        //获取文件的后缀名
        String fileName = file.getOriginalFilename();
        //成功则保存文件至目录
        File newFile = new File(path + fileName);
        //读excel
        readExcelFile(file.getInputStream(), fileName);
        //拷贝文件，性能高效，比原先的方便
        file.transferTo(newFile);

        return "ok";
    }

    /**
     * 读excel操作
     *
     * @param inputStream
     * @param fileName
     */
    public void readExcelFile(InputStream inputStream, String fileName) {

        /**
         * 这个inputStream文件可以来源于本地文件的流，
         *  也可以来源与上传上来的文件的流，也就是MultipartFile的流，
         *  使用getInputStream()方法进行获取。
         */

        /**
         * 然后再读取文件的时候，应该excel文件的后缀名在不同的版本中对应的解析类不一样
         * 要对fileName进行后缀的解析
         */
        Workbook workbook = null;
        try {
            //判断什么类型文件
            if (fileName.endsWith(".xls")) {
                workbook = new HSSFWorkbook(inputStream);
            } else if (fileName.endsWith(".xlsx")) {
                workbook = new XSSFWorkbook(inputStream);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        if (workbook == null) {
            return;
        } else {
            //获取所有的工作表的的数量
            int numOfSheet = workbook.getNumberOfSheets();
            System.out.println(numOfSheet + "--->numOfSheet");
            //遍历表
            for (int i = 0; i < numOfSheet; i++) {
                //获取一个sheet也就是一个工作本。
                Sheet sheet = workbook.getSheetAt(i);
                if (sheet == null) continue;
                //获取一个sheet有多少Row
                int lastRowNum = sheet.getLastRowNum();
                if (lastRowNum == 0) continue;
                Row row;
                for (int j = 1; j <= lastRowNum; j++) {
                    row = sheet.getRow(j);
                    if (row == null) {
                        continue;
                    }
                    //获取一个Row有多少Cell
                    short lastCellNum = row.getLastCellNum();
                    for (int k = 0; k <= lastCellNum; k++) {
                        if (row.getCell(k) == null) {
                            continue;
                        }
                        row.getCell(k).setCellType(Cell.CELL_TYPE_STRING);
                        String res = row.getCell(k).getStringCellValue().trim();
                        //打印出cell(单元格的内容)
                        System.out.println(res);
                    }
                }
            }
        }
    }

}
