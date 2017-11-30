package org.poihelper.sheet;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

public class ColumnDescriptor {
    public boolean boldHeader;
    public String header;

    public boolean bold;
    public String dataFormat;
    public HorizontalAlignment hAlign;
    public VerticalAlignment vAlign;

    public boolean isBoldHeader() {
        return boldHeader;
    }

    public void setBoldHeader(boolean boldHeader) {
        this.boldHeader = boldHeader;
    }

    public String getHeader() {
        return header;
    }

    public void setHeader(String header) {
        this.header = header;
    }

    public boolean isBold() {
        return bold;
    }

    public void setBold(boolean bold) {
        this.bold = bold;
    }

    public String getDataFormat() {
        return dataFormat;
    }

    public void setDataFormat(String dataFormat) {
        this.dataFormat = dataFormat;
    }

    public HorizontalAlignment getHAlign() {
        return hAlign;
    }

    public void setHAlign(HorizontalAlignment hAlign) {
        this.hAlign = hAlign;
    }

    public VerticalAlignment getVAlign() {
        return vAlign;
    }

    public void setVAlign(VerticalAlignment vAlign) {
        this.vAlign = vAlign;
    }


}
