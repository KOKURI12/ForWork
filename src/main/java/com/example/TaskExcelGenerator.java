package com.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.common.usermodel.HyperlinkType;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;

public class TaskExcelGenerator {
    public static void main(String[] args) {
        Workbook workbook = new XSSFWorkbook();
        String safeSheetName = WorkbookUtil.createSafeSheetName("タスク管理");
        Sheet sheet = workbook.createSheet(safeSheetName);

        String[] headers = {
            "タスクID", "タスク名", "WBS番号",
            "ステータス", "優先度", "締切日",
            "担当者", "指示者", "Teams指示リンク", "備考"
        };

        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font boldFont = workbook.createFont();
        boldFont.setBold(true);
        headerStyle.setFont(boldFont);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);

        CellStyle highlightStyle = workbook.createCellStyle();
        highlightStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        highlightStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        Font redFont = workbook.createFont();
        redFont.setColor(IndexedColors.RED.getIndex());

        CellStyle redTextStyle = workbook.createCellStyle();
        redTextStyle.setFont(redFont);

        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
        }

        DataValidationHelper helper = sheet.getDataValidationHelper();
        String[] statusOptions = {"未開始", "進行中", "完了", "確認待ち"};
        CellRangeAddressList statusRange = new CellRangeAddressList(1, 100, 3, 3);
        DataValidationConstraint statusConstraint = helper.createExplicitListConstraint(statusOptions);
        DataValidation statusValidation = helper.createValidation(statusConstraint, statusRange);
        sheet.addValidationData(statusValidation);

        String[] priorityOptions = {"高", "中", "低"};
        CellRangeAddressList priorityRange = new CellRangeAddressList(1, 100, 4, 4);
        DataValidationConstraint priorityConstraint = helper.createExplicitListConstraint(priorityOptions);
        DataValidation priorityValidation = helper.createValidation(priorityConstraint, priorityRange);
        sheet.addValidationData(priorityValidation);

        Row row1 = sheet.createRow(1);
        row1.createCell(0).setCellValue("T-001");
        row1.createCell(1).setCellValue("ログインバグ修正");
        row1.createCell(2).setCellValue("WBS-20250418");
        row1.createCell(3).setCellValue("進行中");
        row1.createCell(4).setCellValue("高");

        Cell dueDateCell = row1.createCell(5);
        Calendar cal = Calendar.getInstance();
        cal.add(Calendar.DATE, 2);
        dueDateCell.setCellValue(cal.getTime());
        dueDateCell.setCellStyle(highlightStyle);

        row1.createCell(6).setCellValue("小林");
        row1.createCell(7).setCellValue("藤田PM");

        CreationHelper creationHelper = workbook.getCreationHelper();
        Cell linkCell = row1.createCell(8);
        linkCell.setCellValue("Teams指示を開く");
        Hyperlink link = creationHelper.createHyperlink(HyperlinkType.URL);
        link.setAddress("https://teams.microsoft.com/example/link-to-message");
        linkCell.setHyperlink(link);
        linkCell.setCellStyle(redTextStyle);

        row1.createCell(9).setCellValue("最優先で対応");

        for (int i = 0; i < headers.length; i++) {
            sheet.autoSizeColumn(i);
        }

        try (FileOutputStream fileOut = new FileOutputStream("タスク管理テンプレート_日本語版.xlsx")) {
            workbook.write(fileOut);
            System.out.println("✅ タスク管理テンプレート_日本語版.xlsx が生成されました！");
        } catch (IOException e) {
            System.out.println("❌ 書き込みエラー：" + e.getMessage());
        }

        try {
            workbook.close();
        } catch (IOException e) {
            System.out.println("❌ クローズ失敗：" + e.getMessage());
        }
    }
}
