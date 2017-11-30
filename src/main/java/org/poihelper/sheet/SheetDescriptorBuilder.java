package org.poihelper.sheet;

import java.util.Arrays;
import java.util.Comparator;
import java.util.List;
import java.util.Optional;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;

public class SheetDescriptorBuilder {
    private static final int INITIAL_ROW_CELL_VALUES = 1;

    private SheetDescriptor sheetDescriptor;

    public static SheetDescriptorBuilder create() {
        return new SheetDescriptorBuilder();
    }

    private SheetDescriptorBuilder() {
        this.sheetDescriptor = new SheetDescriptor();
        sheetDescriptor.setCellValues(Maps.newHashMap());
    }

    public SheetDescriptorBuilder sheet(String sheetName) {
        sheetDescriptor.setSheetName(sheetName);

        return this;
    }

    public SheetDescriptorBuilder rowValues(Object... cellValuesP) {
        return rowValues(getNextRow(), cellValuesP);
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

    private SheetDescriptorBuilder rowValues(Integer rowNum, Object... cellValuesP) {
        sheetDescriptor.getCellValues().put(rowNum, Arrays.asList(cellValuesP));

        return this;
    }

    public SheetDescriptorBuilder columnDescriptors(Object ... columnDescs) {
        List<ColumnDescriptor> columnDescriptors = Lists.newArrayList();

        for (Object object : columnDescs) {
            ColumnDescriptor cd;

            if (object instanceof String) {
                cd = ColumnDescriptorBuilder.create()
                        .header((String) object)
                        .build();
            } else if (object instanceof ColumnDescriptor) {
                cd = (ColumnDescriptor) object;
            } else if (object instanceof ColumnDescriptorBuilder) {
                ColumnDescriptorBuilder builder = (ColumnDescriptorBuilder) object;
                cd = builder.build();
            } else {
                throw new IllegalArgumentException("Invalid type for columnDescriptor");
            }

            columnDescriptors.add(cd);
        }

        sheetDescriptor.setColumnDescriptors(columnDescriptors);

        return this;
    }

    public SheetDescriptor build() {
        return sheetDescriptor;
    }

}
