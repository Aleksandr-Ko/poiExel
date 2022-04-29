package poi.exel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

public class CellStyleAlignmentExample {
    public static void main(String... args) {

        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet("testing");
        sheet.setColumnWidth(0, 10000);
        Row row = sheet.createRow(0);
        row.setHeightInPoints(100);
        Cell cell = row.createCell(0);
        cell.setCellValue("SimpleSolution.dev");

        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.LEFT);
        cellStyle.setVerticalAlignment(VerticalAlignment.TOP);
        cell.setCellStyle(cellStyle);

        try(OutputStream outputStream = new FileOutputStream("sample-left-top-alignment.xls")) {
            workbook.write(outputStream);
        } catch(IOException ex) {
            ex.printStackTrace();
        }
    }
}
