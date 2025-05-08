package com.example.deliverynotesummary.service;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Slf4j
@Service
public class ExcelService {

    public byte[] mergeExcelFiles(List<MultipartFile> files) throws IOException {
        if (files == null || files.isEmpty()) {
            throw new IllegalArgumentException("请选择要合并的Excel文件");
        }

        // 创建新的工作簿用于存储合并后的数据
        Workbook mergedWorkbook = new XSSFWorkbook();
        Sheet mergedSheet = mergedWorkbook.createSheet("合并结果");
        int currentRow = 0;

        // 处理每个上传的文件
        for (int fileIndex = 0; fileIndex < files.size(); fileIndex++) {
            MultipartFile file = files.get(fileIndex);
            Workbook workbook = WorkbookFactory.create(file.getInputStream());
            Sheet sheet = workbook.getSheetAt(0);

            // 如果是第一个文件，复制表头
            if (fileIndex == 0) {
                Row headerRow = sheet.getRow(0);
                if (headerRow != null) {
                    Row newHeaderRow = mergedSheet.createRow(currentRow++);
                    copyRow(headerRow, newHeaderRow);
                }
            }

            // 复制数据行
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row currentRowData = sheet.getRow(i);
                if (currentRowData != null) {
                    Row newRow = mergedSheet.createRow(currentRow++);
                    copyRow(currentRowData, newRow);
                }
            }

            workbook.close();
        }

        // 将合并后的工作簿转换为字节数组
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        mergedWorkbook.write(outputStream);
        mergedWorkbook.close();

        return outputStream.toByteArray();
    }

    private void copyRow(Row source, Row target) {
        for (int i = 0; i < source.getLastCellNum(); i++) {
            Cell sourceCell = source.getCell(i);
            if (sourceCell != null) {
                Cell targetCell = target.createCell(i);
                copyCell(sourceCell, targetCell);
            }
        }
    }

    private void copyCell(Cell source, Cell target) {
        CellStyle newStyle = target.getSheet().getWorkbook().createCellStyle();
        newStyle.cloneStyleFrom(source.getCellStyle());
        target.setCellStyle(newStyle);

        switch (source.getCellType()) {
            case STRING:
                target.setCellValue(source.getStringCellValue());
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(source)) {
                    target.setCellValue(source.getDateCellValue());
                } else {
                    target.setCellValue(source.getNumericCellValue());
                }
                break;
            case BOOLEAN:
                target.setCellValue(source.getBooleanCellValue());
                break;
            case FORMULA:
                target.setCellValue(source.getCellFormula());
                break;
            case BLANK:
                target.setBlank();
                break;
            default:
                target.setBlank();
        }
    }
}