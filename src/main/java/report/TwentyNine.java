package report;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PropertyTemplate;

import java.io.FileOutputStream;
import java.io.IOException;

public class TwentyNine {

    public static void main(String[] args) throws IOException {

        TwentyNine tw = new TwentyNine();
        tw.report("1.xls");

    }

    // формирование таблицы
    public void report(String file) throws IOException {

        Workbook book = new HSSFWorkbook();

        Sheet sheet = book.createSheet("ОПИ ДЛЯ ОКС");
        sheet.setMargin(Sheet.LeftMargin, 0.4 /* inches */);
        sheet.setMargin(Sheet.RightMargin, 0.4 /* inches */);

        // https://overcoder.net/q/677606/определите-необходимую-ориентацию-печати-с-помощью-apache-poi
        PrintSetup printSetup = sheet.getPrintSetup();
        printSetup.setLandscape(true);

        // https://coderoad.ru/18637394/как-установить-строки-заголовка-печати-на-листе-excel-с-помощью-POI
        sheet.setRepeatingRows(CellRangeAddress.valueOf("6"));


        // Заполнить заголовок
        header(book, sheet, fontBasic(book));
        nameColumn(book, sheet, fontBasic(book));
//            body( book, sheet);


        PropertyTemplate propertyTemplate = new PropertyTemplate();
        propertyTemplate.drawBorders(new CellRangeAddress(1, 3, 0, 13),
                BorderStyle.THIN, BorderExtent.ALL);
        propertyTemplate.drawBorders(new CellRangeAddress(5, 7, 0, 13),
                BorderStyle.THIN, BorderExtent.ALL);
        propertyTemplate.applyBorders(sheet);

        // Сохранение документа
        book.write(new FileOutputStream(file));

        System.out.println("\nМожно пробовать");
    }

    private void header(Workbook book, Sheet sheet, Font fontBasic) {
        // Стиль заголовка
        CellStyle baseStyle = styleCellFont(book,fontBasic);

        Row row0 = sheet.createRow(0);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 13));
        initCell(row0.createCell(0), "ФОРМА ЕЖЕДНЕВНОЙ СВОДКИ О ПОСТАВКЕ ОПИ ДЛЯ ОКС", baseStyle);

        Row row1 = sheet.createRow(1);
        sheet.addMergedRegionUnsafe(new CellRangeAddress(1, 1, 0, 13));
        initCell(row1.createCell(0), "Инвестпроект", baseStyle);

        Row row2 = sheet.createRow(2);
        sheet.addMergedRegionUnsafe(new CellRangeAddress(2, 2, 0, 6));
        initCell(row2.createCell(0), "Наименование",baseStyle);
        sheet.addMergedRegionUnsafe(new CellRangeAddress(2, 2, 7, 13));
        initCell(row2.createCell(7), "Код (шифр)",baseStyle);

        Row row3 = sheet.createRow(3);
        sheet.addMergedRegionUnsafe(new CellRangeAddress(3, 3, 0, 6));
        initCell(row3.createCell(0), " ",baseStyle);                                                         // нужно уточнить что сюда вставлять?
        sheet.addMergedRegionUnsafe(new CellRangeAddress(3, 3, 7, 13));
        initCell(row3.createCell(7), " ",baseStyle);                                                         // нужно уточнить что сюда вставлять?
    }

    private void nameColumn(Workbook book, Sheet sheet, Font fontBasic) {

        CellStyle horStyle = styleCellFont(book, fontBasic);

        CellStyle verStyle = styleCellFont(book, fontBasic);
        verStyle.setRotation((short) 90);

        Row row6 = sheet.createRow(5);
        Row row7 = sheet.createRow(6);
//        row2.setHeight((short) 1600);

        Row row8 = sheet.createRow(7);
        for (int i = 1; i < 15; i++) {
            initCell(row8.createCell(i-1), i, horStyle);
        }
    }

    private CellStyle styleCellFont(Workbook book, Font font) {
        CellStyle baseStyle = book.createCellStyle();
        baseStyle.setWrapText(true);                                                                                    // перенос текста
        baseStyle.setAlignment(HorizontalAlignment.CENTER);
        baseStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        baseStyle.setFont(font);
        return baseStyle;
    }

    private Font fontBasic(Workbook book){
        Font font = book.createFont();
        font.setBold(true);
        font.setFontName("Arial");
        font.setFontHeight((short) 180);                                                                                //размер шрифта -> 9 = (9 / (1/20))

        return font;
    }

    private void initCell(Cell cell, String val, CellStyle style) {
        cell.setCellValue(val);
        cell.setCellStyle(style);
    }
    private void initCell(Cell cell, int val, CellStyle style) {
        cell.setCellValue(val);
        cell.setCellStyle(style);
    }

}
