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
        sheet.getHeader().setCenter(HSSFHeader.font("Arial", "bold") + HSSFHeader.fontSize((short) 12) + "СВОДНЫЙ ОТЧЕТ");
        // задаем отступ от края листа для печати
        sheet.setMargin(Sheet.LeftMargin, 0.4);
        sheet.setMargin(Sheet.RightMargin, 0.4);
        // устанавливаем ориентацию листа для печати (альбомная)
        sheet.getPrintSetup().setLandscape(true);
        // выравнивание по центру листа
        sheet.setHorizontallyCenter(true);
        // перенос рядов на каждый лист
        sheet.setRepeatingRows(CellRangeAddress.valueOf("7"));

        // наполнение документа
        header(book, sheet);
        nameColumn(book, sheet);
        quarryData(book, sheet);


        // отображение границ таблицы
        PropertyTemplate propertyTemplate = new PropertyTemplate();
        propertyTemplate.drawBorders(new CellRangeAddress(0, 2, 17, 24),
                BorderStyle.THIN, BorderExtent.ALL);
        // TODO -> заменить рисование границ на всю область данных
        propertyTemplate.drawBorders(new CellRangeAddress(4, 18, 0, 24),
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
        row3.setHeight((short) (2 * 255));

        sheet.addMergedRegion(new CellRangeAddress(0, 2, 0, 1));
        initCell(row0.createCell(0), "Агент:", boldElevenStyle);
        sheet.addMergedRegion(new CellRangeAddress(0, 2, 2, 10));
        initCell(row0.createCell(2), "Тут будет вставка организации", regularStyleLeft);
        sheet.addMergedRegion(new CellRangeAddress(0, 1, 17, 18));
        initCell(row0.createCell(17), "Номер\nдокумента", regularStyle);
        sheet.addMergedRegion(new CellRangeAddress(0, 1, 19, 20));
        initCell(row0.createCell(19), "Дата\nсоставления", regularStyle);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 21, 24));
        initCell(row0.createCell(21), "Отчетный период", regularStyle);

        sheet.addMergedRegion(new CellRangeAddress(1, 1, 21, 22));
        initCell(row1.createCell(21), "c", regularStyle);
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 23, 24));
        initCell(row1.createCell(22), "по", regularStyle);

        sheet.addMergedRegion(new CellRangeAddress(2, 2, 17, 18));
        initCell(row2.createCell(17), "?", regularStyle);
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 19, 20));
        initCell(row2.createCell(19), "?", regularStyle);
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 21, 22));
        initCell(row2.createCell(21), "?", regularStyle);
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 23, 24));
        initCell(row2.createCell(23), "?", regularStyle);
    }

    private void nameColumn(Workbook book, Sheet sheet) {
        CellStyle regularStyle = cellStyle(book, fontRegular(book));
        CellStyle verStyle = cellStyle(book, fontRegular(book));
        verStyle.setRotation((short)90);

        int weightCell = 9;

        Row row4 = sheet.createRow(4);
        Row row5 = sheet.createRow(5);
        row4.setHeight((short) (3 * 255));

        sheet.addMergedRegionUnsafe(new CellRangeAddress(4, 5, 0, 0));
        initCellWidth(sheet,4, row4.createCell(0), "№\nп/п", regularStyle);

        sheet.addMergedRegionUnsafe(new CellRangeAddress(4, 4, 1, 2));
        initCell(row4.createCell(1), "Стройка", regularStyle);
        initCellWidth(sheet,weightCell, row5.createCell(1), "Код", verStyle);
        initCellWidth(sheet,weightCell, row5.createCell(2), "Наименование ", verStyle);

        sheet.addMergedRegionUnsafe(new CellRangeAddress(4, 4, 3, 4));
        initCell(row4.createCell(3), "Объект", regularStyle);
        initCellWidth(sheet,weightCell, row5.createCell(3), "Код", verStyle);
        initCellWidth(sheet,weightCell, row5.createCell(4), "Наименование", verStyle);

        sheet.addMergedRegionUnsafe(new CellRangeAddress(4, 4, 5, 6));
        initCell(row4.createCell(5), "Контрагент\n(подрядчик СМР)", regularStyle);
        initCellWidth(sheet,weightCell, row5.createCell(5), "Код", verStyle);
        initCellWidth(sheet,weightCell, row5.createCell(6), "Наименование", verStyle);

        sheet.addMergedRegionUnsafe(new CellRangeAddress(4, 4, 7, 8));
        initCell(row4.createCell(7), "Договор\nподряда", regularStyle);
        initCellWidth(sheet,weightCell, row5.createCell(7), "Номер", verStyle);
        initCellWidth(sheet,weightCell, row5.createCell(8), "Дата", verStyle);

        sheet.addMergedRegionUnsafe(new CellRangeAddress(4, 4, 9, 10));
        initCell(row4.createCell(9), "Карьер", regularStyle);
        initCellWidth(sheet,weightCell, row5.createCell(9), "Код", verStyle);
        initCellWidth(sheet,weightCell, row5.createCell(10), "Наименование карьера", verStyle);

        sheet.addMergedRegionUnsafe(new CellRangeAddress(4, 5, 11, 11));
        initCellWidth(sheet,weightCell, row4.createCell(11), "Наименование материала (ОПИ)", verStyle);

        sheet.addMergedRegionUnsafe(new CellRangeAddress(4, 5, 12, 12));
        initCellWidth(sheet,weightCell, row4.createCell(12), "Номенклатурный номер", verStyle);

        sheet.addMergedRegionUnsafe(new CellRangeAddress(4, 4, 13, 14));
        initCell(row4.createCell(13), "Единица измерения", regularStyle);
        initCellWidth(sheet,weightCell, row5.createCell(13), "Код", verStyle);
        initCellWidth(sheet,weightCell, row5.createCell(14), "Наименование", verStyle);

        sheet.addMergedRegionUnsafe(new CellRangeAddress(4, 5, 15, 15));
        initCellWidth(sheet,12, row4.createCell(15),
                "Остаток ОПИ (в основном состоянии*)\nна начало отчетного периода", verStyle);

        sheet.addMergedRegionUnsafe(new CellRangeAddress(4, 5, 16, 16));
        initCellWidth(sheet,12, row4.createCell(16),
                "Получено ОПИ (в основном состоянии)\nза отчетный месяц", verStyle);

        sheet.addMergedRegionUnsafe(new CellRangeAddress(4, 5, 17, 17));
        initCellWidth(sheet,weightCell, row4.createCell(17), "№ накладной", verStyle);

        sheet.addMergedRegionUnsafe(new CellRangeAddress(4, 5, 18, 18));
        initCellWidth(sheet,weightCell, row4.createCell(18), "Дата накладной", verStyle);

        sheet.addMergedRegionUnsafe(new CellRangeAddress(4, 4, 19, 21));
        initCell(row4.createCell(19), "Израсходовано ОПИ\n(в основном состоянии)\nза отчетный месяц", regularStyle);
        initCellWidth(sheet,weightCell, row5.createCell(19), "количество ОПИ", verStyle);
        initCellWidth(sheet,weightCell, row5.createCell(20), "№, дата Справки о движении\nОПИ к Отчету ф.№М-29", verStyle);
        initCellWidth(sheet,weightCell, row5.createCell(21), "№, дата акта ф.№КС-2", verStyle);

        sheet.addMergedRegionUnsafe(new CellRangeAddress(4, 4, 22, 23));
        initCell(row4.createCell(22), "СПРАВОЧНО:\n"
                + "объем использованных ОПИ исходя из проектного объема работ", regularStyle);

        initCellWidth(sheet,16, row5.createCell(22), "израсходовано ОПИ исходя\nиз проектного объема работ\n"
                + "за отчетный период\n(в соответствии с ф.№КС-2)", verStyle);
        initCellWidth(sheet,16, row5.createCell(23), "коэффициент (соотношение)\nобъема ОПИ в основном\n"
                + "состоянии и ОПИ исходя из\nпроектного объема работ", verStyle);

        sheet.addMergedRegionUnsafe(new CellRangeAddress(4, 5, 24, 24));
        initCellWidth(sheet,weightCell, row4.createCell(24), "Остаток ОПИ (в основном состоянии)\nна конец отчетного периода", verStyle);

        Row row6 = sheet.createRow(6);
        for (int i = 1; i < 26; i++) {
            initCell(row6.createCell(i - 1), String.valueOf(i), regularStyle);
        }
    }

    private void quarryData(Workbook book, Sheet sheet){
        CellStyle regularStyle = cellStyle(book, fontRegular(book));
        CellStyle verStyle = cellStyle(book, fontRegular(book));
        verStyle.setRotation((short)90);

        Row row7 = sheet.createRow(7);
        initCell(row7.createCell(0), "", regularStyle);
        for (int i = 1; i < 26; i++) {
            initCell(row7.createCell(i - 1), "", verStyle);
        }

        Row row8 = sheet.createRow(8);
        sheet.addMergedRegion(new CellRangeAddress(8,8,0,14));
        cellNameTotal(regularStyle, verStyle, row8);

        Row row9 = sheet.createRow(9);
        Row row10 = sheet.createRow(10);
        rowMet(sheet, regularStyle, verStyle, row9, row10,"по сторонним\nкарьерам");

        Row row11 = sheet.createRow(11);
        Row row12 = sheet.createRow(12);
        rowMet(sheet, regularStyle, verStyle, row11, row12, "по\nсубъектам");

        Row row13 = sheet.createRow(13);
        Row row14 = sheet.createRow(14);
        rowMet(sheet, regularStyle, verStyle, row13, row14, "по\nпоставщикам");

        Row row15 = sheet.createRow(15);
        Row row16 = sheet.createRow(16);
        rowMet(sheet, regularStyle, verStyle, row15, row16, "по\nподрядчикам");

        Row row17 = sheet.createRow(17);
        Row row18 = sheet.createRow(18);
        rowMet(sheet, regularStyle, verStyle, row17, row18, "по\nобъектам");
    }

    private void rowMet(Sheet sheet, CellStyle cellRegularStyle, CellStyle cellVerStyle, Row rowName, Row rowTotal, String text) {
        sheet.addMergedRegion(new CellRangeAddress(rowName.getRowNum(),rowName.getRowNum(),0,2));
        sheet.addMergedRegion(new CellRangeAddress(rowTotal.getRowNum(),rowTotal.getRowNum(),0,14));
        rowName.setHeight((short)(2*255));
        initCell(rowName.createCell(0), text, cellRegularStyle);
        cellNameTotal(cellRegularStyle, cellVerStyle, rowTotal);
    }

    private void cellNameTotal(CellStyle regularStyle, CellStyle verStyle, Row rowTotal) {
        initCell(rowTotal.createCell(0), "Итого", regularStyle);
        for (int i = 15; i < 26; i++) {
            if(i == 16 | i == 17 | i == 20 | i == 23 | i == 25){
                initCell(rowTotal.createCell(i - 1), "х", regularStyle);
            }else {
                initCell(rowTotal.createCell(i - 1), "", verStyle);
            }
        }
    }

    private void initCell(Cell cell, String text, CellStyle cellStyle) {
        cell.setCellValue(text);
        cell.setCellStyle(cellStyle);
    }
    private void initCellWidth(Sheet sheet, int wight, Cell cell, String text, CellStyle cellStyle) {
        sheet.setColumnWidth(cell.getColumnIndex(), wight*256);
        cell.setCellValue(text);
        cell.setCellStyle(cellStyle);
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
}