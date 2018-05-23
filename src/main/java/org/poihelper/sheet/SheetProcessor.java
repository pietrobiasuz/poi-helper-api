package org.poihelper.sheet;

import java.math.BigDecimal;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.common.collect.Lists;

public class SheetProcessor {
    private static final int ROW_HEADER_VALUES = 0;
    private static final int INITIAL_COLUMN = 0;

    private Logger logger = LoggerFactory.getLogger(SheetProcessor.class);

    public void process(Workbook wb, SheetDescriptor sheetDescriptor) {
        List<CellStyle> columnStyles = createColumnStyles(wb, sheetDescriptor.getColumnDescriptors());

        Sheet sheet = wb.createSheet(sheetDescriptor.getSheetName());
        createHeader(wb, sheet, sheetDescriptor.getColumnDescriptors());
        createCells(sheet, sheetDescriptor.getCellValues(), columnStyles);
    }

    private List<CellStyle> createColumnStyles(Workbook wb, List<ColumnDescriptor> columnDescriptors) {
        List<CellStyle> columnStyles = Lists.newArrayList();

        for (ColumnDescriptor columnDescriptor : columnDescriptors) {
            CellStyle columnStyle = wb.createCellStyle();

            if (columnDescriptor.getHAlign() != null) {
                columnStyle.setAlignment(columnDescriptor.getHAlign());
            }

            if (columnDescriptor.getVAlign() != null) {
                columnStyle.setVerticalAlignment(columnDescriptor.getVAlign());
            }

            if (columnDescriptor.getDataFormat() != null
                            && !columnDescriptor.getDataFormat().isEmpty()) {
                columnStyle.setDataFormat(wb.createDataFormat().getFormat(columnDescriptor.getDataFormat()));
            }

            if (columnDescriptor.isBold()) {
                Font boldFont = createBoldFont(wb);
                columnStyle.setFont(boldFont);
            }

            logger.info("created cell style {} for workbook {}", columnStyle, wb);

            columnStyles.add(columnStyle);
        }

        return columnStyles;
    }

    private Font createBoldFont(Workbook wb) {
        Font font = wb.createFont();
        logger.info("created font {} on wb {}", font, wb);

        font.setBold(true);

        return font;
    }

    private void createHeader(Workbook wb, Sheet sheet, List<ColumnDescriptor> columnDescriptors) {
        Integer cellNum = INITIAL_COLUMN;
        Row row = sheet.createRow(ROW_HEADER_VALUES);

        Font boldFont = createBoldFont(wb);
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setFont(boldFont);

        for (ColumnDescriptor columnDesc : columnDescriptors) {
            Cell cell = row.createCell(cellNum++);

            if (columnDesc.isBoldHeader()) {
                cell.setCellStyle(cellStyle);
            }

            cell.setCellValue(columnDesc.getHeader());
        }
    }

    private void createCells(Sheet sheet, Map<Integer, List<Object>> cellValues, List<CellStyle> columnStyles) {
        for (Entry<Integer, List<Object>> rowCells : cellValues.entrySet()) {
            Integer rowNum = rowCells.getKey();

            Row row = sheet.createRow(rowNum);

            logger.info("Processing row {}", rowNum);
            createRowCells(columnStyles, rowCells, row);
        }
    }

    private void createRowCells(List<CellStyle> columnStyles, Entry<Integer, List<Object>> rowCells, Row row) {
        int columnNum = INITIAL_COLUMN;
        for (Object cellValue : rowCells.getValue()) {
            logger.info("Processing column {}", columnNum);

            Cell cell = row.createCell(columnNum);

            CellStyle columnStyle = columnStyles.get(columnNum);
            cell.setCellStyle(columnStyle);

            assignCellValue(cell, cellValue);

            columnNum++;
        }
    }

    private void assignCellValue(Cell cell, Object cellValue) {
        if (cellValue instanceof Calendar) {
            cell.setCellValue((Calendar) cellValue);
            return;
        }

        if (cellValue instanceof Date) {
            cell.setCellValue((Date) cellValue);
            return;
        }

        if (cellValue instanceof Integer) {
            cell.setCellValue((Integer) cellValue);
            return;
        }

        if (cellValue instanceof Double) {
            cell.setCellValue((Double) cellValue);
            return;
        }

        if (cellValue instanceof BigDecimal) {
            BigDecimal bd = (BigDecimal) cellValue;
            cell.setCellValue(bd.doubleValue());
            return;
        }

        if (cellValue instanceof Boolean) {
            cell.setCellValue((Boolean) cellValue);
            return;
        }

        if (cellValue != null) {
            String cellValueStr = cellValue.toString();
            cell.setCellValue(cellValueStr);
        }
    }

}
