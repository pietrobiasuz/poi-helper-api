package org.poihelper.sheet.cellstyle;

import java.util.List;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.common.collect.Lists;

/**
 * Componente para processar e criar as {@link CellStyle} da planilha.
 *
 * @author pietro.biasuz
 */
public class CellStyleProcessor {
    private Logger logger = LoggerFactory.getLogger(CellStyleProcessor.class);

    //evitar duplicacao do estilo bold font pois geralmente eh muito utilizado
    private CellStyle cellStyleBoldFont;

    private Workbook wb;

    public CellStyleProcessor(Workbook wb) {
        this.wb = wb;
    }

    public List<CellStyle> createColumnCellStyles(int numberOfColumns, List<CellStyleDescriptor> cellStyleDescriptors) {
        List<CellStyle> columnStyles = Lists.newArrayList();

        for (int ind = 0; ind < numberOfColumns; ind++) {
            CellStyle columnStyle;

            if (cellStyleDescriptors.size() > ind) {
                CellStyleDescriptor cellStyleDescriptor = cellStyleDescriptors.get(ind);

                columnStyle = processCellStyleDescriptor(cellStyleDescriptor);
                logger.trace("Created cell style based on column index {} descriptor", ind);
            } else {
                columnStyle = wb.createCellStyle();
                logger.trace("Created default cell style for column index {}", ind);
            }

            columnStyles.add(columnStyle);
        }

        return columnStyles;
    }

    public CellStyle processCellStyleDescriptor(CellStyleDescriptor columnDescriptor) {
        CellStyle columnStyle;

        if (columnDescriptor.isBold()) {
            columnStyle = getCellStyleBoldFont(wb);
        } else {
            columnStyle = wb.createCellStyle();
        }

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

        return columnStyle;
    }

    private CellStyle getCellStyleBoldFont(Workbook wb) {
        if (cellStyleBoldFont == null) {
            this.cellStyleBoldFont = generateCellStyleBoldFont(wb);
        }

        return this.cellStyleBoldFont;
    }

    private CellStyle generateCellStyleBoldFont(Workbook wb) {
        Font boldFont = createBoldFont();
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setFont(boldFont);

        return cellStyle;
    }

    private Font createBoldFont() {
        Font font = wb.createFont();
        logger.trace("Created font {} on wb {}", font, wb);

        font.setBold(true);

        return font;
    }

}
