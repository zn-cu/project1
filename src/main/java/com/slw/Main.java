package com.slw;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.*;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.*;
/*
* 1.1先用Java读取自己电脑中所有文件
* (不确定用户是什么系统，为方便跨平台工作，读取的是用户目录下的documents文件夹)
* (部分文件夹没有权限，直接略过，顺便打印出来哪些文件夹访问不了)
* 1.2按照文件名称，大小，更新时间分别排序
* 1.3输出包含文件全路径的三个Excel的sheet，
* 1.4在项目下创建自己的分支，将Java代码和Excel文件上传到分支上
* */
public class Main {

    public static void main(String[] args) throws IOException {
        //获取路径
        String userHome = System.getProperty("user.home");
        Path startPath = Paths.get(userHome, "documents");
        //创建列表存储文件对象
        List<File> files = new ArrayList<>();
        //递归遍历文件
        Files.walkFileTree(startPath, new SimpleFileVisitor<Path>() {
            //正常文件就添加列表里就行了
            public FileVisitResult visitFile(Path file, BasicFileAttributes attrs) throws IOException {
                if (attrs.isRegularFile()) {
                    files.add(file.toFile());
                }
                return FileVisitResult.CONTINUE;
            }

            //权限不足或出现异常的就直接跳过
            public FileVisitResult visitFileFailed(Path file, IOException exc) throws IOException {
                if (exc instanceof AccessDeniedException) {
                    System.err.println("Permission denied for: " + file.toString());
                    return FileVisitResult.CONTINUE; // 忽略访问被拒绝的文件/目录
                }
                throw exc; // 如果是其他类型的异常，则重新抛出
            }
        });
        //先排序再将列表写进excel中
        sortAndWrite(files, "output.xlsx");
    }

    private static void sortAndWrite(List<File> files, String fileName) throws IOException {
        //创建新的工作簿
        Workbook workbook = new XSSFWorkbook();
        //按名称，大小，最后修改时间排序
        for (String sheetName : Arrays.asList("ByName", "BySize", "ByLastModified")) {
            Sheet sheet = workbook.createSheet(sheetName);
            List<File> sortedFiles = new ArrayList<>(files);
            switch (sheetName) {
                case "ByName":
                    sortedFiles.sort(Comparator.comparing(File::getName));
                    break;
                case "BySize":
                    sortedFiles.sort(Comparator.comparingLong(File::length));
                    break;
                case "ByLastModified":
                    sortedFiles.sort(Comparator.comparingLong(File::lastModified));
                    break;
            }
            int rowNum = 0;
            //写入工作簿
            for (File file : sortedFiles) {
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(file.getAbsolutePath());
                row.createCell(1).setCellValue(file.length());
                row.createCell(2).setCellValue(file.lastModified());
            }
        }
        //将工作簿写入excel文件
        try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
            workbook.write(fileOut);
        }
        workbook.close();
    }
}