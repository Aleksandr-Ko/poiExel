package poi.exel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

public class BorderStyleThinExample {
    public static void main(String... args) {

        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet("sample");    // создаем лист

        sheet.setColumnWidth(1, 6000);           // устанавливаем ширину колонки выбраной ячейки

        Row row = sheet.createRow(1);                     // берем ряд номер
        row.setHeightInPoints(100);                               // устанавливаем высоту ряда

        Cell cell = row.createCell(1);                    // берем конкретную ячейку
        cell.setCellValue("SimpleSolution.dev");                 // заносим данные

        // Стиль ячейки с границами и цветом границы
        CellStyle style = workbook.createCellStyle();
        style.setBorderBottom(BorderStyle.THICK);
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
