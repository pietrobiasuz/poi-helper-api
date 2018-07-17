package org.poihelper.sheet;

import java.util.Arrays;
import java.util.List;
import java.util.Set;

import org.apache.poi.ss.usermodel.CellStyle;
import org.poihelper.sheet.cellstyle.CellStyleDescriptor;
import org.poihelper.sheet.cellstyle.CellStyleDescriptorBuilder;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;

public class SheetDescriptorBuilder {
    private static final int DEFAULT_HEADER_LINES = 1;
    private static final String DEFAULT_SHEET_NAME = "Sheet-1";

    private SheetDescriptor sheetDescriptor;
    private Integer nextRow = -1;

    public static SheetDescriptorBuilder create() {
        return new SheetDescriptorBuilder();
    }

    private SheetDescriptorBuilder() {
        this.sheetDescriptor = new SheetDescriptor();

        sheetDescriptor.setCellValues(Maps.newHashMap());
        sheetDescriptor.setColumnCellStyleDescriptors(Maps.newHashMap());
        sheetDescriptor.setHeaderLines(DEFAULT_HEADER_LINES);
        sheetDescriptor.setHeaderStyle(CellStyleDescriptorBuilder.create()
                        .bold(true)
                        .build());
        sheetDescriptor.setSheetName(DEFAULT_SHEET_NAME);
    }

    public SheetDescriptorBuilder sheet(String sheetName) {
        sheetDescriptor.setSheetName(sheetName);

        return this;
    }

    /**
     * @deprecated Use {@link SheetDescriptorBuilder#nextRowValues(Object...)} instead.
     *
     * @param cellValuesP
     * @return
     */
    @Deprecated
    public SheetDescriptorBuilder rowValues(Object... cellValuesP) {
        return rowValues(getNextRow(), cellValuesP);
    }

    public SheetDescriptorBuilder nextRowValues(Object... cellValuesP) {
        return rowValues(getNextRow(), cellValuesP);
    }

    public SheetDescriptorBuilder continueRowValues(Object... cellValuesP) {
        return rowValues(getLastRow(), cellValuesP);
    }

    private Integer getNextRow() {
        return ++nextRow;
    }

    private Integer getLastRow() {
        return nextRow;
    }

    private SheetDescriptorBuilder rowValues(Integer rowNum, Object ... cellValuesP) {
        List<Object> values = sheetDescriptor.getCellValues().containsKey(rowNum) ?
                        sheetDescriptor.getCellValues().get(rowNum) :  Lists.newArrayList();

        values.addAll(Arrays.asList(cellValuesP));

        sheetDescriptor.getCellValues().put(rowNum, values);

        return this;
    }

    /**
     * Seta a quantidade de linhas de cabecalho da planilha.
     *
     * @param headerLines quantidade de linhas
     */
    public SheetDescriptorBuilder headerLines(int headerLines) {
        this.sheetDescriptor.setHeaderLines(headerLines);

        return this;
    }

    /**
     * Seta o {@link CellStyle} para o cabecalho da planilha.
     *
     * <p>
     * Sobre-escreve o cellStyle da coluna.
     *
     * @param cellStyleDescriptor
     */
    public SheetDescriptorBuilder headerStyle(CellStyleDescriptor cellStyleDescriptor) {
        this.sheetDescriptor.setHeaderStyle(cellStyleDescriptor);

        return this;
    }

    /**
     * Seta o {@link CellStyle} para a coluna da planilha.
     *
     * @param colNum non-zero-based da coluna
     */
    public SheetDescriptorBuilder columnCellStyle(int colNum, CellStyleDescriptor cellStyle) {
        sheetDescriptor.getColumnCellStyleDescriptors().put(Integer.valueOf(adjustNum(colNum)), cellStyle);

        return this;
    }

    /**
     * See {@link SheetDescriptorBuilder#columnCellStyle(int, CellStyleDescriptor)}
     */
    public SheetDescriptorBuilder columnCellStyle(Set<Integer> colNums, CellStyleDescriptor cellStyle) {
        colNums.parallelStream()
            .forEach(colNum -> columnCellStyle(colNum, cellStyle));

        return this;
    }

    /**
     * Seta os {@link CellStyle} para as colunas da planilha em ordem.
     *
     * @param cellStyleDescriptors descriptors na ordem das colunas
     */
    public SheetDescriptorBuilder columnCellStyles(CellStyleDescriptor ... cellStyleDescriptors) {
        List<CellStyleDescriptor> cellStyles = Arrays.asList(cellStyleDescriptors);

        for (int ind = 0; ind < cellStyles.size(); ind++) {
            columnCellStyle(ind, cellStyles.get(ind));
        }

        return this;
    }

    private int adjustNum(int nonZeroBased) {
        if (nonZeroBased < 1) {
            throw new IllegalArgumentException("Column and row numbers are non zero based. Must be grater than ZERO.");
        }

        return nonZeroBased - 1;
    }

    public SheetDescriptor build() {
         return sheetDescriptor;
    }

}
