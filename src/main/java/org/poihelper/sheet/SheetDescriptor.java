package org.poihelper.sheet;

import java.util.List;
import java.util.Map;

public class SheetDescriptor {
    private String sheetName;
    private Map<Integer, List<Object>> cellValues;
    private List<ColumnDescriptor> columnDescriptors;

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public Map<Integer, List<Object>> getCellValues() {
        return cellValues;
    }

    public void setCellValues(Map<Integer, List<Object>> cellValues) {
        this.cellValues = cellValues;
    }

    public List<ColumnDescriptor> getColumnDescriptors() {
        return columnDescriptors;
    }

    public void setColumnDescriptors(List<ColumnDescriptor> columnDescriptors) {
        this.columnDescriptors = columnDescriptors;
    }

}
