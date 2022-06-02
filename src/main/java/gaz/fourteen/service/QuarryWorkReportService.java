package gaz.fourteen.service;

import gaz.fourteen.dto.QuarryWorkDTO;
import org.apache.poi.hssf.usermodel.HSSFHeader;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PropertyTemplate;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class QuarryWorkReportService {
    public static void main(String[] args) throws IOException {

        QuarryWorkDTO dto = new QuarryWorkDTO();

        QuarryWorkDataService data = new QuarryWorkDataService();
        data.getQuarryWorkDTO(dto);

        QuarryWorkReportService fourteen14 = new QuarryWorkReportService();
        fourteen14.createReport("14.xls", dto);
    }

    // формирование таблицы
    public void createReport(String file, QuarryWorkDTO dto) throws IOException {
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
        nameColumn(book, sheet, dto);
        addRow(book, sheet, dto);


// отображение границ таблицы
        PropertyTemplate propertyTemplate = new PropertyTemplate();
        propertyTemplate.drawBorders(new CellRangeAddress(0, 2, 0, 9), BorderStyle.THIN, BorderExtent.ALL);
        propertyTemplate.drawBorders(new CellRangeAddress(4, sheet.getLastRowNum(), 0, 9), BorderStyle.THIN, BorderExtent.ALL);
        propertyTemplate.applyBorders(sheet);

// Сохранение документа
        FileOutputStream out = new FileOutputStream(file);
        book.write(out);
        out.close();

        System.out.println("\nМожно пробовать\n14.xls");
    }
    // наименование колонок основной таблицы
    private void nameColumn(Workbook book, Sheet sheet, QuarryWorkDTO dto) {
// стиль ячейки
        CellStyle cellStyle = cellStyle(book, font(book));

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

        initCell(row0.createCell(0), "Инвестпроект", cellStyle);
        initCell(row1.createCell(0), "Наименование", cellStyle);
        initCell(row1.createCell(5), "Код (шифр)", cellStyle);
        initCell(row2.createCell(0), dto.getNameInvestProject(), cellStyle);
        initCell(row2.createCell(5), dto.getCodeInvestProject(), cellStyle);
        initCell(row4.createCell(0), "Субподрядный договор на добычу ОПИ, с видом работ «добыча ОПИ на карьере»", cellStyle);
        initCell(row4.createCell(3), "Карьер", cellStyle);
        initCell(row4.createCell(5), "Сводка", cellStyle);

        int cellWightOne = 8;
        int cellWightTwo = 10;
        int cellWightThree = 14;

        initCellWidth(sheet, cellWightThree, row4.createCell(2), "Контрагент\n(наименование\nорганизации)", cellStyle);
        initCellWidth(sheet, cellWightTwo + 1, row5.createCell(0), "Номер\nдоговора", cellStyle);
        initCellWidth(sheet, cellWightTwo + 1, row5.createCell(1), "Дата\nдоговора", cellStyle);
        initCellWidth(sheet, cellWightOne, row5.createCell(3), "Субъект\nРФ", cellStyle);
        initCellWidth(sheet, cellWightThree, row5.createCell(4), "Наименование\nкарьера", cellStyle);
        initCellWidth(sheet, cellWightThree, row5.createCell(5), "Дата, за\nкоторую\nподается\nсводка", cellStyle);
        initCellWidth(sheet, cellWightThree + 1, row5.createCell(6), "Вид работ на\nкарьере", cellStyle);
        initCellWidth(sheet, cellWightThree + 1, row5.createCell(7), "ОПИ (Материал)", cellStyle);
        initCellWidth(sheet, cellWightTwo, row5.createCell(8), "Объем\nвсего,\nкуб.м", cellStyle);
        initCellWidth(sheet, cellWightThree + 1, row5.createCell(9), "Объем всего, тн", cellStyle);

// нумеруем колонки под наименованием
        for (int i = 1; i < 11; i++) {
            initCell(row6.createCell(i - 1), String.valueOf(i), cellStyle);
        }
    }

    // вставка строк с данными и итогом
    private void addRow(Workbook book, Sheet sheet, QuarryWorkDTO dto) {

        CellStyle cellStyle = cellStyle(book, font(book));

        addSetData(sheet, dto.getDummy(), cellStyle);
    }

    // стиль шрифта
    private Font font(Workbook book) {
        Font font = book.createFont();
        font.setFontName("Times New Roman");
        font.setFontHeight((short) (9 * 20));
        return font;
    }

    // стиль ячейки
    private CellStyle cellStyle(Workbook book, Font font) {
        CellStyle style = book.createCellStyle();
        style.setWrapText(true);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFont(font);
        return style;
    }

    // метод добавление текста в ячейку
    private void initCell(Cell cell, String val, CellStyle style) {
        cell.setCellValue(val);
        cell.setCellStyle(style);
    }

    // метод добавление текста в ячейку с указанием ширины колонки
    private void initCellWidth(Sheet sheet, int wightColum, Cell cell, String text, CellStyle cellStyle) {
        sheet.setColumnWidth(cell.getColumnIndex(), wightColum * 255);
        cell.setCellValue(text);
        cell.setCellStyle(cellStyle);
    }

    // метод добавление ряда стройки
    private void addSetData(Sheet sheet, List<String> list, CellStyle cellStyle) {

        Row row = sheet.createRow(sheet.getLastRowNum() + 1);

        for (int i = 0; i < list.size(); i++) {
            initCell(row.createCell(i), list.get(i), cellStyle);
        }
    }
}
