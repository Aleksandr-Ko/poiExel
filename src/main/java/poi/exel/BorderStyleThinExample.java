package poi.exel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

public class BorderStyleThinExample {
    public static void main(String... args) {
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet("sample");
        sheet.setColumnWidth(1, 6000);
        Row row = sheet.createRow(1);
        row.setHeightInPoints(100);
        Cell cell = row.createCell(1);
        cell.setCellValue("SimpleSolution.dev");

        // Стиль ячейки с границами и цветом границы
        CellStyle style = workbook.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        cell.setCellStyle(style);

        try (OutputStream fileOut = new FileOutputStream("sample-border-thin-black.xls")) {
            workbook.write(fileOut);
            workbook.close();
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }
}
