package gaz.eleven.service;

import gaz.eleven.dto.ZDAgentExpenseDTO;
import org.apache.poi.hssf.usermodel.HSSFHeader;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PropertyTemplate;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class ZDAgentExpenseReportService {
    public static void main(String[] args) throws IOException {
        ZDAgentExpenseDTO dto = new ZDAgentExpenseDTO();

        ZDAgentExpenseDataService data = new ZDAgentExpenseDataService();
        data.getZDAgentExpenseDTO(dto);

        ZDAgentExpenseReportService el = new ZDAgentExpenseReportService();
        el.createReport("11.xls", dto);

    }

    public void createReport(String file,ZDAgentExpenseDTO dto) throws IOException {
        // создание документа
        Workbook book = new HSSFWorkbook();
        // создаём лист в документе
        Sheet sheet = book.createSheet("3Д Отчет агента");

        // верхний колонтитул
        // TODO -> возможно это лишняя запись
        sheet.getHeader().setCenter(HSSFHeader.font("Arial", "bold") + HSSFHeader.fontSize((short) 12)
                + " 3Д ОТЧЕТ АГЕНТА\n"
                + " «О РАСХОДЕ ОСНОВНЫХ МАТЕРИАЛОВ В СТРОИТЕЛЬСТВЕ В СОПОСТАВЛЕНИИ С РАСХОДОМ,\n"
                + " ОПРЕДЕЛЕННЫМ ПО ПРОИЗВОДСТВЕННЫМ НОРМАМ» ");

        // задаем отступ от края листа для печати
        sheet.setMargin(Sheet.TopMargin, 1.4);
        sheet.setMargin(Sheet.LeftMargin, 0.4);
        sheet.setMargin(Sheet.RightMargin, 0.4);
        // устанавливаем ориентацию листа для печати (альбомная)
        sheet.getPrintSetup().setLandscape(true);
        // выравнивание по центру листа
        sheet.setHorizontallyCenter(true);
        // перенос рядов на каждый лист
        sheet.setRepeatingRows(CellRangeAddress.valueOf("3"));

        // наполнение документа
        nameColumn(book, sheet);
        filling(book, dto);

        // отображение границ таблицы
        PropertyTemplate propertyTemplate = new PropertyTemplate();
        // TODO -> заменить рисование границ на всю область данных
        propertyTemplate.drawBorders(new CellRangeAddress(0, sheet.getLastRowNum(), 0, 9),BorderStyle.THIN, BorderExtent.ALL);
        propertyTemplate.applyBorders(sheet);

        // Сохранение документа
        book.write(new FileOutputStream(file));

        System.out.println("\nМожно пробовать\n11.xls");
    }

    // создание шапки таблицы
    private void nameColumn(Workbook book, Sheet sheet) {

        CellStyle cellStyleBold = cellStyle(book, font(book, true, 10));
        CellStyle cellStyleBoldSmall = cellStyle(book, font(book, true, 8));

        Row row0 = sheet.createRow(0);
        Row row1 = sheet.createRow(1);
        Row row2 = sheet.createRow(2);

        row0.setHeight((short) (2 * 255));
        row1.setHeight((short) (3 * 255));

        sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 0));
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 1, 3));

        for (int i = 4; i < 10; i++) {
            sheet.addMergedRegion(new CellRangeAddress(0, 1, i, i));
        }

        initCell(row0.createCell(1), "Первичный документ", cellStyleBold);

        initCellWidth(sheet, 4, row0.createCell(0), "№\nп/п", cellStyleBold);
        initCellWidth(sheet, 16, row1.createCell(1), "Наименование", cellStyleBold);
        initCellWidth(sheet, 10, row1.createCell(2), "Номер", cellStyleBold);
        initCellWidth(sheet, 8, row1.createCell(3), "Дата", cellStyleBold);

        initCellWidth(sheet, 16, row0.createCell(4), "Наименование\nматериала", cellStyleBold);
        initCellWidth(sheet, 20, row0.createCell(5), "Номенклатурный\nномер", cellStyleBold);
        initCellWidth(sheet, 12, row0.createCell(6), "Остаток\nна начало\nмесяца", cellStyleBold);
        initCellWidth(sheet, 12, row0.createCell(7), "Поступило\nв текущем\nмесяце", cellStyleBold);
        initCellWidth(sheet, 16, row0.createCell(8), "Использовано\nв текущем\nмесяце", cellStyleBold);
        initCellWidth(sheet, 12, row0.createCell(9), "Остаток\nна конец\nмесяца", cellStyleBold);

        for (int i = 1; i < 11; i++) {
            initCell(row2.createCell(i - 1), String.valueOf(i), cellStyleBoldSmall);
        }
    }

    // заполнение таблицы данными
    private void filling(Workbook book, ZDAgentExpenseDTO dto) {

        CellStyle styleCell = cellStyle(book, font(book, false, 8));

        Sheet sheet = book.getSheet("3Д Отчет агента");
        sheet.addMergedRegion(new CellRangeAddress(4, 4, 0, 5));

        Row row3 = sheet.createRow(3);
        Row row4 = sheet.createRow(4);

        initCell(row4.createCell(0), "ИТОГО:", styleCell);

        List<String> dummy = dto.getDummy();
        List<String> total = dto.getTotal();

        for (int i = 0; i < dummy.size(); i++) {
            initCell(row3.createCell(i), dummy.get(i), styleCell);
        }

        for (int i = 0; i < total.size(); i++) {
            initCell(row4.createCell(i + 6), total.get(i), styleCell);
        }
    }

    // стиль шрифта
    private Font font(Workbook book, boolean bold, int fontSize) {
        Font font = book.createFont();
        font.setFontName("Arial");
        font.setBold(bold);
        font.setFontHeight((short) (fontSize * 20));
        return font;
    }

    // стиль ячейки
    private CellStyle cellStyle(Workbook book, Font font) {
        CellStyle style = book.createCellStyle();
        style.setWrapText(true);                                                                                        // перенос текста
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFont(font);
        return style;
    }

    // добавление текста в ячейку
    private void initCell(Cell cell, String text, CellStyle cellStyle) {
        cell.setCellValue(text);
        cell.setCellStyle(cellStyle);
    }

    // добавление текста в ячейку с указанием ширины колонки
    private void initCellWidth(Sheet sheet, int wightColum, Cell cell, String text, CellStyle cellStyle) {
        sheet.setColumnWidth(cell.getColumnIndex(), wightColum * 255);
        cell.setCellValue(text);
        cell.setCellStyle(cellStyle);
    }
}
