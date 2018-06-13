package org.poihelper.sheet.cellstyle;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

/**
 * Descritor para o estilo de uma celula da planilha.
 *
 * @author pietro.biasuz
 */
public class CellStyleDescriptor {

    public boolean bold;
    public String dataFormat;
    public HorizontalAlignment hAlign;
    public VerticalAlignment vAlign;

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
