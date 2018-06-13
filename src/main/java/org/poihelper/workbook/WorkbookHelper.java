package org.poihelper.workbook;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.poihelper.sheet.SheetDescriptor;
import org.poihelper.sheet.SheetDescriptorProcessor;

/**
 * Helper para criar workbook a partir de uma lista de {@link SheetDescriptor}.
 *
 * @author pietro.biasuz
 */
public class WorkbookHelper {
    public static final String MS_EXCEL_TYPE = "application/vnd.ms-excel";

    private SheetDescriptorProcessor sheetProcessor = new SheetDescriptorProcessor();

    public ByteArrayOutputStream create(List<SheetDescriptor> sheetDescriptors) throws IOException {
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();

        try (Workbook wb = new XSSFWorkbook()) {
            sheetDescriptors.stream().forEach(sd -> sheetProcessor.process(wb, sd));

            wb.write(byteArrayOutputStream);
        }

        return byteArrayOutputStream;
    }

}
