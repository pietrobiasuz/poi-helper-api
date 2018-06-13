package org.poihelper.sheet.cellstyle;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

public class CellStyleDescriptorBuilder {
    private CellStyleDescriptor cellStyleDescriptor;

    public static CellStyleDescriptorBuilder create() {
        return new CellStyleDescriptorBuilder();
    }

    private CellStyleDescriptorBuilder() {
        this.cellStyleDescriptor = new CellStyleDescriptor();
        this.cellStyleDescriptor.setBold(false);
    }

    public CellStyleDescriptorBuilder bold(boolean bold) {
        this.cellStyleDescriptor.setBold(bold);

        return this;
    }

    public CellStyleDescriptorBuilder dataFormat(String dataFormat) {
        this.cellStyleDescriptor.setDataFormat(dataFormat);

        return this;
    }

    public CellStyleDescriptorBuilder vAlign(VerticalAlignment vAlign) {
        this.cellStyleDescriptor.setVAlign(vAlign);

        return this;
    }

    public CellStyleDescriptorBuilder hAlign(HorizontalAlignment hAlign) {
        this.cellStyleDescriptor.setHAlign(hAlign);

        return this;
    }

    public CellStyleDescriptor build() {
        return this.cellStyleDescriptor;
    }

}
