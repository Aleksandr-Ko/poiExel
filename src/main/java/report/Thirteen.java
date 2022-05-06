package report;

import org.apache.poi.hssf.usermodel.HSSFHeader;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PropertyTemplate;

import java.io.FileOutputStream;
import java.io.IOException;

public class Thirteen {
    public static void main(String[] args) throws IOException {


        Thirteen tr = new Thirteen();
        tr.createReport("13.xls");

    }

    public void createReport(String file) throws IOException {
        // создание документа
        Workbook book = new HSSFWorkbook();
        // создаём лист в документе
        Sheet sheet = book.createSheet("Сводный отчет");
        // верхний колонтитул
        sheet.getHeader().setCenter(HSSFHeader.font("Arial", "bold") + HSSFHeader.fontSize((short) 12) + "ФОРМА СВОДНОГО ОТЧЕТА");
        // задаем отступ от края листа для печати
        sheet.setMargin(Sheet.LeftMargin, 0.4);
        sheet.setMargin(Sheet.RightMargin, 0.4);
        // устанавливаем ориентацию листа для печати (альбомная)
        sheet.getPrintSetup().setLandscape(true);
        // выравнивание по центру листа
        sheet.setHorizontallyCenter(true);

        // наполнение документа
        header(book,sheet);


        // отображение границ таблицы
        PropertyTemplate propertyTemplate = new PropertyTemplate();
        propertyTemplate.drawBorders(new CellRangeAddress(1, 3, 17, 24),
                BorderStyle.THIN, BorderExtent.ALL);
        propertyTemplate.applyBorders(sheet);

        // Сохранение документа
        book.write(new FileOutputStream(file));

        System.out.println("\nМожно пробовать\n13.xls");
    }

    private void header(Workbook book, Sheet sheet) {
        // стили для ячеек

        CellStyle twelveStyle = cellStyle(book, fontTwelve(book));
        CellStyle boldElevenStyle = cellStyle(book, fontBoldEleven(book));
        CellStyle regularStyle = cellStyle(book, fontRegular(book));
        CellStyle regularStyleLeft = cellStyle(book, fontRegular(book));
        regularStyleLeft.setAlignment(HorizontalAlignment.LEFT);

        Row row0 = sheet.createRow(0);
        Row row1 = sheet.createRow(1);
        Row row2 = sheet.createRow(2);
        Row row3 = sheet.createRow(3);

        sheet.addMergedRegionUnsafe(new CellRangeAddress(0, 0, 0, 24));
        initCell(row0.createCell(0), "Сводный отчет", twelveStyle);

        sheet.addMergedRegion(new CellRangeAddress(1,3,0,0));
        initCell(row1.createCell(0), "Агент:", boldElevenStyle);
        sheet.addMergedRegion(new CellRangeAddress(1,3,1,8));

        initCell(row1.createCell(1), "Тут будет вставка организации", regularStyleLeft);
        sheet.addMergedRegion(new CellRangeAddress(1,2,17,18));
        initCell(row1.createCell(17), "Номер\nдокумента", regularStyle);
        sheet.addMergedRegion(new CellRangeAddress(1,2,19,20));
        initCell(row1.createCell(19), "Дата\nсоставления", regularStyle);
        sheet.addMergedRegion(new CellRangeAddress(1,1,21,24));
        initCell(row1.createCell(21), "Отчетный период", regularStyle);
        sheet.addMergedRegion(new CellRangeAddress(2,2,21,22));

        initCell(row2.createCell(21), "c", regularStyle);
        sheet.addMergedRegion(new CellRangeAddress(2,2,23,24));
        initCell(row2.createCell(21), "по", regularStyle);

        sheet.addMergedRegion(new CellRangeAddress(3,3,17,18));
        initCell(row3.createCell(17), "?", regularStyle);
        sheet.addMergedRegion(new CellRangeAddress(3,3,19,20));
        initCell(row3.createCell(19), "?", regularStyle);
        sheet.addMergedRegion(new CellRangeAddress(3,3,21,22));
        initCell(row3.createCell(21), "?", regularStyle);
        sheet.addMergedRegion(new CellRangeAddress(3,3,23,24));
        initCell(row3.createCell(23), "?", regularStyle);

        Row row4 = sheet.createRow(4);
        row4.setHeight((short) (2*256));
        sheet.addMergedRegionUnsafe(new CellRangeAddress(4, 4, 0, 24));

    }

    private Font fontRegular(Workbook book) {
        Font font = book.createFont();
        font.setFontName("Times New Roman");
        font.setFontHeight((short) (10 * 20));
        return font;
    }
    private Font fontBoldEleven(Workbook book) {
        Font font = book.createFont();
        font.setFontName("Times New Roman");
        font.setBold(true);
        font.setFontHeight((short) (11 * 20));
        return font;
    }
    private Font fontEleven(Workbook book) {
        Font font = book.createFont();
        font.setFontName("Times New Roman");
        font.setFontHeight((short) (11 * 20));
        return font;
    }
    private Font fontTwelve(Workbook book) {
        Font font = book.createFont();
        font.setFontName("Times New Roman");
        font.setFontHeight((short) (12 * 20));
        return font;
    }

    private CellStyle cellStyle(Workbook book, Font font) {
        CellStyle style = book.createCellStyle();
        style.setWrapText(true);                                                                                        // перенос текста
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFont(font);
        return style;
    }

    private void initCell(Cell cell, String val, CellStyle style) {
        cell.setCellValue(val);
        cell.setCellStyle(style);
    }

}
