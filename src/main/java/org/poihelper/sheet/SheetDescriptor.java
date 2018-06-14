package org.poihelper.sheet;

import java.util.List;

import java.util.Map;

import org.poihelper.sheet.cellstyle.CellStyleDescriptor;

/**
 * Descriptor de uma sheet de um workbook.
 *
 * @author pietro.biasuz
 */
public class SheetDescriptor {
    private String sheetName;

    private int headerLines;
    private CellStyleDescriptor headerStyle;

    private Map<Integer, List<Object>> cellValues;
    private Map<Integer, CellStyleDescriptor> columnCellStyleDescriptors;

    /**
     * @return nome da planilha (ira na aba do workbook)
     */
    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    /**
     * @return Map com key com o numero da linha (inclui cabecalho) e com value com os respectivos valores
     */
    public Map<Integer, List<Object>> getCellValues() {
        return cellValues;
    }

    public void setCellValues(Map<Integer, List<Object>> cellValues) {
        this.cellValues = cellValues;
    }

    /**
     * @return numero de linhas que farao parte do cabecalho da planilha
     */
    public int getHeaderLines() {
        return headerLines;
    }

    public void setHeaderLines(int headerLines) {
        this.headerLines = headerLines;
    }

    /**
     * @return descritor do estilo de todas as celulas do cabecalho
     */
    public CellStyleDescriptor getHeaderStyle() {
        return headerStyle;
    }

    public void setHeaderStyle(CellStyleDescriptor headerStyle) {
        this.headerStyle = headerStyle;
    }

    public Map<Integer, CellStyleDescriptor> getColumnCellStyleDescriptors() {
        return columnCellStyleDescriptors;
    }

    public void setColumnCellStyleDescriptors(Map<Integer, CellStyleDescriptor> columnCellStyleDescriptors) {
        this.columnCellStyleDescriptors = columnCellStyleDescriptors;
    }

}
