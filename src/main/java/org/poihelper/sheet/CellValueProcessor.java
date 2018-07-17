package org.poihelper.sheet;

import java.math.BigDecimal;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;

/**
 * Componente para processar o valor das celulas da planilha
 *
 * @author pietro.biasuz
 */
public class CellValueProcessor {

    public static void assignCellValue(Cell cell, Object cellValue) {
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
