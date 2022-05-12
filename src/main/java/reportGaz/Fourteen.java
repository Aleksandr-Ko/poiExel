package reportGaz;

import org.apache.poi.hssf.usermodel.HSSFHeader;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PropertyTemplate;

import java.io.FileOutputStream;
import java.io.IOException;

public class Fourteen {
    public static void main(String[] args) throws IOException {

        Fourteen fourteen = new Fourteen();
        fourteen.createReport("14.xls");
    }

    // формирование таблицы
    public void createReport(String file) throws IOException {
        // создание документа
        Workbook book = new HSSFWorkbook();
        // создаём лист в документе
        Sheet sheet = book.createSheet("Сводка работ на карьере");

        // верхний колонтитул
        // TODO: возможно это лишняя запись
        sheet.getHeader().setCenter(HSSFHeader.font("Arial", "bold")
                + HSSFHeader.fontSize((short) 12) + "ЕЖЕДНЕВНАЯ СВОДКА РЕЗУЛЬТАТОВ РАБОТ НА КАРЬЕРЕ");

        // задаем отступ от края листа для печати
        sheet.setMargin(Sheet.BottomMargin, 0.4);
        sheet.setMargin(Sheet.LeftMargin, 0.4);
        sheet.setMargin(Sheet.RightMargin, 0.4);
        // устанавливаем ориентацию листа для печати (альбомная)
        sheet.getPrintSetup().setLandscape(true);
        // выравнивание по центру листа
        sheet.setHorizontallyCenter(true);
        // перенос рядов на каждый лист
        sheet.setRepeatingRows(CellRangeAddress.valueOf("7"));

        // Наполнение документа
        nameColumn(book, sheet, fontStyle(book));
//        dataColumn(book, sheet, fontStyle(book));

        // отображение границ таблицы
        PropertyTemplate propertyTemplate = new PropertyTemplate();
        propertyTemplate.drawBorders(new CellRangeAddress(0, 2, 0, 9),
                BorderStyle.THIN, BorderExtent.ALL);
        // TODO -> заменить рисование границ на всю область данных
        propertyTemplate.drawBorders(new CellRangeAddress(4, 6, 0, 9),
                BorderStyle.THIN, BorderExtent.ALL);
        propertyTemplate.applyBorders(sheet);

        // Сохранение документа
        book.write(new FileOutputStream(file));

        System.out.println("\nМожно пробовать\n14.xls");
    }

    private void nameColumn(Workbook book, Sheet sheet, Font font) {
        // стиль ячейки
        CellStyle horStyle = cellStyle(book, font);

        Row row0 = sheet.createRow(0);
        Row row1 = sheet.createRow(1);
        Row row2 = sheet.createRow(2);
        Row row3 = sheet.createRow(3);
        Row row4 = sheet.createRow(4);
        Row row5 = sheet.createRow(5);
        Row row6 = sheet.createRow(6);

        row3.setHeight((short) (2 * 256));
        row4.setHeight((short) (4 * 256));
        row5.setHeight((short) (5 * 228));

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 9));
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 4));
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 5, 9));
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 0, 4));
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 5, 9));

        sheet.addMergedRegion(new CellRangeAddress(3, 3, 0, 9));
        sheet.addMergedRegion(new CellRangeAddress(4, 4, 0, 1));
        sheet.addMergedRegion(new CellRangeAddress(4, 5, 2, 2));
        sheet.addMergedRegion(new CellRangeAddress(4, 4, 3, 4));
        sheet.addMergedRegion(new CellRangeAddress(4, 4, 5, 9));

        initCell(row0.createCell(0), "Инвестпроект", horStyle);
        initCell(row1.createCell(0), "Наименование", horStyle);
        initCell(row1.createCell(5), "Код (шифр)", horStyle);
        initCell(row2.createCell(0), "", horStyle);                                                          // нужно уточнить что сюда вставлять?
        initCell(row2.createCell(7), "", horStyle);                                                          // нужно уточнить что сюда вставлять?

        initCell(row4.createCell(0), "Субподрядный договор на добычу ОПИ, с видом работ «добыча ОПИ на карьере»", horStyle);
        initCell(row4.createCell(3), "Карьер", horStyle);
        initCell(row4.createCell(5), "Сводка", horStyle);

        int cellWightOne = 8;
        int cellWightTwo = 10;
        int cellWightThree = 14;

        initCellWidth(sheet, cellWightThree, row4.createCell(2), "Контрагент\n(наименование\nорганизации)", horStyle);
        initCellWidth(sheet, cellWightTwo + 1, row5.createCell(0), "Номер\nдоговора", horStyle);
        initCellWidth(sheet, cellWightTwo + 1, row5.createCell(1), "Дата\nдоговора", horStyle);
        initCellWidth(sheet, cellWightOne, row5.createCell(3), "Субъект\nРФ", horStyle);
        initCellWidth(sheet, cellWightThree, row5.createCell(4), "Наименование\nкарьера", horStyle);
        initCellWidth(sheet, cellWightThree, row5.createCell(5), "Дата, за\nкоторую\nподается\nсводка", horStyle);
        initCellWidth(sheet, cellWightThree + 1, row5.createCell(6), "Вид работ на\nкарьере", horStyle);
        initCellWidth(sheet, cellWightThree + 1, row5.createCell(7), "ОПИ (Материал)", horStyle);
        initCellWidth(sheet, cellWightTwo, row5.createCell(8), "Объем\nвсего,\nкуб.м", horStyle);
        initCellWidth(sheet, cellWightThree + 1, row5.createCell(9), "Объем всего, тн", horStyle);

        // нумеруем колонки под наименованием
        for (int i = 1; i < 11; i++) {
            initCell(row6.createCell(i - 1), String.valueOf(i), horStyle);
        }
    }

    private Font fontStyle(Workbook book) {
        Font font = book.createFont();
        font.setBold(true);
        font.setFontName("Arial");
        font.setFontHeight((short) (9 * 20));                                                                                //размер шрифта -> 9 = (9 / (1/20))
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

    private void initCellWidth(Sheet sheet, int wightColum, Cell cell, String text, CellStyle cellStyle) {
        sheet.setColumnWidth(cell.getColumnIndex(), wightColum * 255);
        cell.setCellValue(text);
        cell.setCellStyle(cellStyle);
    }
}
