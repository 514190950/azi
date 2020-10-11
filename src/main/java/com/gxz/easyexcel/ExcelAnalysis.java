package com.gxz.easyexcel;


import javafx.scene.control.Alert;
import javafx.scene.control.DialogPane;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.filechooser.FileSystemView;
import java.io.*;
import java.nio.channels.FileChannel;
import java.time.*;
import java.util.*;
import java.util.stream.Collectors;

/**
 * @author gxz gongxuanzhang@foxmail.com
 **/
public class ExcelAnalysis {

    public static void analysis(File file) {
        String randomPath = UUID.randomUUID().toString().replaceAll("-", "") + "xlsx";
        File copyExcel = new File(randomPath);
        FileSystemView fsv = FileSystemView.getFileSystemView();
        String desktopPath = fsv.getHomeDirectory().getAbsolutePath();
        copy(file, copyExcel);
        List<Row> weekMap = new ArrayList<>();
        List<Row> monthMap = new ArrayList<>();
        List<Row> threeMonthMap = new ArrayList<>();
        List<Row> halfYearMap = new ArrayList<>();
        List<Row> yearMap = new ArrayList<>();
        List<Row> updateList = new ArrayList<>();
        long timeStamp = System.currentTimeMillis();
        try (XSSFWorkbook newExcel = new XSSFWorkbook(new FileInputStream(copyExcel))) {

            XSSFSheet sheet = newExcel.getSheetAt(5);
            int lastRowNum = sheet.getLastRowNum();
            for (int i = 1; i <= lastRowNum; i++) {
                Row row = sheet.getRow(i);
                Cell date = row.getCell(8);
                LocalDateTime parse = getDateTime(date);
                long l = Duration.between(parse, LocalDateTime.now()).toDays();
                if (l <= 7L) {
                    weekMap.add(row);
                } else if (l <= 30) {
                    monthMap.add(row);
                    Cell cell = row.getCell(10);
                    String stringCellValue = cell.getStringCellValue();
                    if (!stringCellValue.equals("一个月")) {
                        updateList.add(row);
                        cell.setCellValue("一个月");
                    }
                } else if (l <= 90) {
                    threeMonthMap.add(row);
                    Cell cell = row.getCell(10);
                    String stringCellValue = cell.getStringCellValue();
                    if (!stringCellValue.equals("三个月")) {
                        updateList.add(row);
                        cell.setCellValue("三个月");
                    }
                } else if (l <= 180) {
                    halfYearMap.add(row);
                    Cell cell = row.getCell(10);
                    String stringCellValue = cell.getStringCellValue();
                    if (!stringCellValue.equals("半年")) {
                        updateList.add(row);
                        cell.setCellValue("半年");
                    }
                } else {
                    yearMap.add(row);
                    Cell cell = row.getCell(10);
                    String stringCellValue = cell.getStringCellValue();
                    if (!stringCellValue.equals("一年及以上")) {
                        updateList.add(row);
                        cell.setCellValue("一年及以上");
                    }
                }
            }
            moveRow(weekMap, newExcel.getSheetAt(0), "7天");
            moveRow(monthMap, newExcel.getSheetAt(1), "一个月");
            moveRow(threeMonthMap, newExcel.getSheetAt(2), "三个月");
            moveRow(halfYearMap, newExcel.getSheetAt(3), "半年");
            moveRow(yearMap, newExcel.getSheetAt(4), "一年及以上");
            File outputFile = new File(desktopPath + "/" + timeStamp + ".xlsx");
            FileOutputStream fileOutputStream = new FileOutputStream(outputFile);
            newExcel.write(fileOutputStream);
            copyExcel.delete();
            Alert alert = new Alert(Alert.AlertType.INFORMATION);
            alert.setTitle("完事了");
            String contentText = "输出的excel在" + outputFile.getAbsolutePath();

            if (updateList.isEmpty()) {
                alert.setHeaderText("没有人变化");
            } else {
                List<String> collect =
                        updateList.stream().map((row) -> row.getCell(2).getStringCellValue()).collect(Collectors.toList());
                if (collect.size() > 7) {
                    alert.setHeaderText("有变化的人共有" + updateList.size() + "人,人数太多显示不下 具体看文件信息");
                    File updateInfoFile = new File(desktopPath + "/人员名单.txt");
                    try (FileOutputStream fileOutputStream1 = new FileOutputStream(updateInfoFile)) {
                        String info = "有变化的人如下\r\n" + collect.toString();
                        fileOutputStream1.write(info.getBytes());
                    }
                    contentText += "\r\n人员变化具体内容在" + updateInfoFile.getAbsolutePath();
                } else {
                    alert.setHeaderText("有变化的人是" + collect + "\r\n共" + updateList.size() + "人");
                }
            }
            alert.setContentText(contentText);
            alert.showAndWait();
        } catch (Exception ignore) {
            Alert alert = new Alert(Alert.AlertType.ERROR);
            alert.setTitle("完了");
            alert.setContentText("解析错误 请问隐匿大佬哪里出现问题了");
            alert.showAndWait();
        }

    }

    private static LocalDateTime getDateTime(Cell date) {
        String stringCellValue = date.getStringCellValue();
        String year = stringCellValue.substring(0, 4);
        String month = stringCellValue.substring(5, 7);
        String day = stringCellValue.substring(8, 10);
        return LocalDateTime.of(Integer.parseInt(year), Integer.parseInt(month),
                Integer.parseInt(day), 0, 0, 0);
    }

    private static LocalDateTime getDateTime(Row row) {
        return getDateTime(row.getCell(8));
    }


    public static void copy(File file1, File file2) {
        try (FileInputStream fileInputStream = new FileInputStream(file1);
             FileOutputStream fileOutputStream = new FileOutputStream(file2);
             FileChannel fisChannel = fileInputStream.getChannel();
             FileChannel fosChannel = fileOutputStream.getChannel()) {
            fisChannel.transferTo(0, fisChannel.size(), fosChannel);
        } catch (IOException ignore) {
        }
    }


    private static void moveRow(List<Row> list, Sheet sheet, String description) {
        int lastRowNum = sheet.getLastRowNum();
        for (int i = 1; i <= lastRowNum; i++) {
            Row row = sheet.getRow(i);
            sheet.removeRow(row);
        }
        list.sort(Comparator.comparing(ExcelAnalysis::getDateTime));
        for (int i = 1; i < list.size(); i++) {
            Row row = sheet.createRow(i);
            Row value = list.get(i);
            Cell indexCell = row.createCell(0);
            CellStyle cellStyle = value.getCell(1).getCellStyle();
            indexCell.setCellValue(i);
            indexCell.setCellStyle(cellStyle);
            for (int cellIndex = 1; cellIndex <= 8; cellIndex++) {
                Cell next = value.getCell(cellIndex);
                Cell cell = row.createCell(cellIndex);
                cell.setCellType(next.getCellTypeEnum());
                cell.setCellStyle(next.getCellStyle());
                CellType cellTypeEnum = next.getCellTypeEnum();
                switch (cellTypeEnum) {
                    case STRING:
                        cell.setCellValue(next.getStringCellValue());
                        break;
                    case NUMERIC:
                        cell.setCellValue(next.getNumericCellValue());
                        break;
                    default:
                        break;
                }
            }
            Cell dayCell = row.createCell(9);
            dayCell.setCellFormula("DATEDIF(I" + (i + 1) + ",TODAY(),\"D\")");
            dayCell.setCellStyle(cellStyle);
            Cell newCell = row.createCell(10);
            newCell.setCellStyle(cellStyle);
            newCell.setCellValue(description);
        }
    }

    public static void main(String[] args) {
        System.out.println(11);
    }

}
