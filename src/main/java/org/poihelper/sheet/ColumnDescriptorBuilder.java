package org.poihelper.sheet;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

public class ColumnDescriptorBuilder {
    private ColumnDescriptor columnDescriptor;

    public static ColumnDescriptorBuilder create() {
        return new ColumnDescriptorBuilder();
    }

    private ColumnDescriptorBuilder() {
        this.columnDescriptor = createColumnDesc();
    }

    private ColumnDescriptor createColumnDesc() {
        ColumnDescriptor columnDescriptor = new ColumnDescriptor();
        columnDescriptor.setBold(false);
        columnDescriptor.setBoldHeader(true);

        return columnDescriptor;
    }

    public ColumnDescriptorBuilder header(String title) {
        columnDescriptor.setHeader(title);

        return this;
    }

    public ColumnDescriptorBuilder bold(boolean bold) {
        this.columnDescriptor.setBold(bold);

        return this;
    }

    public ColumnDescriptorBuilder dataFormat(String dataFormat) {
        this.columnDescriptor.setDataFormat(dataFormat);

        return this;
    }

    public ColumnDescriptorBuilder vAlign(VerticalAlignment vAlign) {
        this.columnDescriptor.setVAlign(vAlign);

        return this;
    }

    public ColumnDescriptorBuilder hAlign(HorizontalAlignment hAlign) {
        this.columnDescriptor.setHAlign(hAlign);

        return this;
    }

    public ColumnDescriptor build() {
        return this.columnDescriptor;
    }


}
