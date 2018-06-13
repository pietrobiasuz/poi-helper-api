package org.poihelper.sheet;

import java.util.Arrays;
import java.util.Comparator;
import java.util.List;
import java.util.Optional;

import org.poihelper.sheet.cellstyle.CellStyleDescriptor;
import org.poihelper.sheet.cellstyle.CellStyleDescriptorBuilder;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;

public class SheetDescriptorBuilder {
    private static final int INITIAL_ROW_CELL_VALUES = 0;
    private static final int DEFAULT_HEADER_LINES = 1;
    private static final String DEFAULT_SHEET_NAME = "Sheet-1";

    private SheetDescriptor sheetDescriptor;

    public static SheetDescriptorBuilder create() {
        return new SheetDescriptorBuilder();
    }

    private SheetDescriptorBuilder() {
        this.sheetDescriptor = new SheetDescriptor();

        sheetDescriptor.setCellValues(Maps.newHashMap());
        sheetDescriptor.setColumnCellStyleDescriptors(Lists.newArrayList());
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
        Optional<Integer> max = sheetDescriptor.getCellValues().keySet().stream()
                        .max(Comparator.naturalOrder());

        Integer nextRow;

        if (max.isPresent()) {
            nextRow = max.get() + 1;
        } else {
            nextRow = INITIAL_ROW_CELL_VALUES;
        }

        return nextRow;
    }

    private Integer getLastRow() {
        Optional<Integer> max = sheetDescriptor.getCellValues().keySet().stream()
                        .max(Comparator.naturalOrder());

        return max.get();
    }

    private SheetDescriptorBuilder rowValues(Integer rowNum, Object ... cellValuesP) {
        List<Object> values = sheetDescriptor.getCellValues().containsKey(rowNum) ?
                        sheetDescriptor.getCellValues().get(rowNum) :  Lists.newArrayList();

        values.addAll(Arrays.asList(cellValuesP));

        sheetDescriptor.getCellValues().put(rowNum, values);

        return this;
    }

    public SheetDescriptorBuilder headerLines(int headerLines) {
        this.sheetDescriptor.setHeaderLines(headerLines);

        return this;
    }

    public SheetDescriptorBuilder headerStyle(CellStyleDescriptor cellStyleDescriptor) {
        this.sheetDescriptor.setHeaderStyle(cellStyleDescriptor);

        return this;
    }

    public SheetDescriptorBuilder columnDescriptors(Object ... columnDescs) {
        List<CellStyleDescriptor> columnDescriptors = Lists.newArrayList();

        for (Object object : columnDescs) {
            CellStyleDescriptor cd;

            if (object instanceof String) {
                cd = CellStyleDescriptorBuilder.create().build();
            } else if (object instanceof CellStyleDescriptor) {
                cd = (CellStyleDescriptor) object;
            } else if (object instanceof CellStyleDescriptorBuilder) {
                CellStyleDescriptorBuilder builder = (CellStyleDescriptorBuilder) object;
                cd = builder.build();
            } else {
                throw new IllegalArgumentException("Invalid type for columnDescriptor");
            }

            columnDescriptors.add(cd);
        }

        sheetDescriptor.setColumnCellStyleDescriptors(columnDescriptors);

        return this;
    }

    public SheetDescriptor build() {
         return sheetDescriptor;
    }

}
