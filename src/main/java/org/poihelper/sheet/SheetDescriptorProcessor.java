package org.poihelper.sheet;

import java.util.Comparator;
import java.util.List;
import java.util.Map.Entry;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.poihelper.sheet.cellstyle.CellStyleProcessor;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Processador central da planilha.
 *
 * <p>
 * Para as celulas do corpo da planilha serao aplicados o mesmo estilo para todas as celulas da mesma coluna,
 * conforme {@link SheetDescriptor#getColumnCellStyleDescriptors()}.
 *
 * <p>
 * Para as celulas do cabecalho da planilha serao aplicados o estilo em {@link SheetDescriptor#getHeaderStyle()}.
 *
 * @author pietro.biasuz
 */
public class SheetDescriptorProcessor {
    private Logger logger = LoggerFactory.getLogger(SheetDescriptorProcessor.class);

    private static final int INITIAL_COLUMN = 0;
    private static final int ZERO = 0;

    private Workbook wb;
    private SheetDescriptor sheetDescriptor;

    public void process(Workbook wb, SheetDescriptor sheetDescriptor) {
        this.wb = wb;
        this.sheetDescriptor = sheetDescriptor;

        Sheet sheet = wb.createSheet(sheetDescriptor.getSheetName());
        createCells(sheet);
    }

    private int getNumberOfColumns() {
        return sheetDescriptor.getCellValues().entrySet().stream()
            .map(entry -> entry.getValue().size())
            .max(Comparator.naturalOrder())
            .orElse(ZERO);
    }

    private void createCells(Sheet sheet) {
        CellStyleProcessor cellStyleProcessor = new CellStyleProcessor(wb);

        List<CellStyle> columnStyles = cellStyleProcessor.createColumnCellStyles(getNumberOfColumns(),
                        sheetDescriptor.getColumnCellStyleDescriptors());
        CellStyle headerStyle = cellStyleProcessor.processCellStyleDescriptor(sheetDescriptor.getHeaderStyle());

        for (Entry<Integer, List<Object>> rowCells : sheetDescriptor.getCellValues().entrySet()) {
            Integer rowNum = rowCells.getKey();

            Row row = sheet.createRow(rowNum);

            if (isRowInHeader(rowNum)) {
                logger.trace("Processing header row {}", rowNum);

                createHeaderCells(headerStyle, rowCells.getValue(), row);
                continue;
            }

            logger.trace("Processing row {}", rowNum);
            createRowCells(columnStyles, rowCells.getValue(), row);
        }
    }

    private boolean isRowInHeader(int rowNum) {
        return rowNum < sheetDescriptor.getHeaderLines();
    }

    private void createRowCells(List<CellStyle> cellStyles, List<Object> values, Row row) {
        int columnNum = INITIAL_COLUMN;

        for (Object cellValue : values) {
            createCell(cellStyles.get(columnNum), cellValue, columnNum, row);

            columnNum++;
        }
    }

    private void createHeaderCells(CellStyle cellStyle, List<Object> values, Row row) {
        int columnNum = INITIAL_COLUMN;

        for (Object cellValue : values) {
            createCell(cellStyle, cellValue, columnNum, row);

            columnNum++;
        }
    }

    private void createCell(CellStyle cellStyle, Object cellValue, int columnNum, Row row) {
        Cell cell = row.createCell(columnNum);

        cell.setCellStyle(cellStyle);

        CellValueProcessor.assignCellValue(cell, cellValue);
    }

}
