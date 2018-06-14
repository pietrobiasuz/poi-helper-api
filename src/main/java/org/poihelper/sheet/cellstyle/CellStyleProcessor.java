package org.poihelper.sheet.cellstyle;

import java.util.List;
import java.util.Map;

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

    private CellStyle defaultCellStyle;
    //evitar duplicacao do estilo bold font pois geralmente eh muito utilizado
    private CellStyle cellStyleBoldFont;

    private Workbook wb;


    public CellStyleProcessor(Workbook wb) {
        this.wb = wb;
    }

    /**
     * Cria CellStyle para todas as colunas da planilha.
     *
     * <p>
     * Caso nao seja definida uma CellStyleDescriptor para a coluna, utiliza uma CellStyle default.
     *
     * @param numberOfColumns
     * @param map
     * @return
     */
    public List<CellStyle> createColumnCellStyles(int numberOfColumns, Map<Integer, CellStyleDescriptor> map) {
        List<CellStyle> columnStyles = Lists.newArrayList();

        for (int colNum = 0; colNum < numberOfColumns; colNum++) {
            CellStyle columnStyle;

            if (map.containsKey(colNum)) {
                CellStyleDescriptor cellStyleDescriptor = map.get(colNum);

                columnStyle = processCellStyleDescriptor(cellStyleDescriptor);
                logger.trace("Created cell style based on column index {} descriptor", colNum);
            } else {
                columnStyle = wb.createCellStyle();
                logger.trace("Created default cell style for column index {}", colNum);
            }

            columnStyles.add(columnStyle);
        }

        return columnStyles;
    }

    public CellStyle processCellStyleDescriptor(CellStyleDescriptor cellStyleDescriptor) {
        CellStyle columnStyle;

        if (cellStyleDescriptor.isBold()) {
            columnStyle = getCellStyleBoldFont();
        } else {
            columnStyle = getDefaultCellStyle();
        }

        if (cellStyleDescriptor.getHAlign() != null) {
            columnStyle.setAlignment(cellStyleDescriptor.getHAlign());
        }

        if (cellStyleDescriptor.getVAlign() != null) {
            columnStyle.setVerticalAlignment(cellStyleDescriptor.getVAlign());
        }

        if (cellStyleDescriptor.getDataFormat() != null
                        && !cellStyleDescriptor.getDataFormat().isEmpty()) {
            columnStyle.setDataFormat(wb.createDataFormat().getFormat(cellStyleDescriptor.getDataFormat()));
        }

        return columnStyle;
    }

    private CellStyle getCellStyleBoldFont() {
        if (cellStyleBoldFont == null) {
            this.cellStyleBoldFont = generateCellStyleBoldFont(wb);
        }

        return this.cellStyleBoldFont;
    }

    private CellStyle getDefaultCellStyle() {
        if (defaultCellStyle == null) {
            this.defaultCellStyle = wb.createCellStyle();
        }

        return this.defaultCellStyle;
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
