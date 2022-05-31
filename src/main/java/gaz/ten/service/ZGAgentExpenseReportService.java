package gaz.ten.service;

import gaz.ten.dto.ZGAgentExpenseDTO;
import org.apache.poi.hssf.usermodel.HSSFHeader;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PropertyTemplate;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class ZGAgentExpenseReportService {
    public static void main(String[] args) throws IOException {

        ZGAgentExpenseDTO dto = new ZGAgentExpenseDTO();

        ZGAgentExpenseDataService data = new ZGAgentExpenseDataService();
        data.getZGAgentExpenseDTO(dto);

        ZGAgentExpenseReportService tn = new ZGAgentExpenseReportService();
        tn.createReport("10.xls", dto);
    }


    public void createReport(String file, ZGAgentExpenseDTO dto) throws IOException {
        // создание документа
        Workbook book = new HSSFWorkbook();
        // создаём лист в документе
        Sheet sheet = book.createSheet("3Г Отчет агента");

        // верхний колонтитул
        // TODO -> возможно это лишняя запись
        sheet.getHeader().setCenter(HSSFHeader.font("Arial", "bold") + HSSFHeader.fontSize((short) 12)
                + " 3Г ОТЧЕТ АГЕНТА\n"
                + " «О РАСХОДЕ ОСНОВНЫХ МАТЕРИАЛОВ В СТРОИТЕЛЬСТВЕ В СОПОСТАВЛЕНИИ С РАСХОДОМ,\n"
                + " ОПРЕДЕЛЕННЫМ ПО ПРОИЗВОДСТВЕННЫМ НОРМАМ»");

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
        propertyTemplate.drawBorders(new CellRangeAddress(0, sheet.getLastRowNum(), 0, 19),
                BorderStyle.THIN, BorderExtent.ALL);
        propertyTemplate.applyBorders(sheet);

        // Сохранение документа
        book.write(new FileOutputStream(file));

        System.out.println("\nМожно пробовать\n10.xls");
    }

    // создание шапки таблицы
    private void nameColumn(Workbook book, Sheet sheet) {

        CellStyle cellHorStyle = cellStyle(book, font(book, true, 8));
        CellStyle cellVerStyle = cellStyle(book, font(book, true, 8));
        cellVerStyle.setRotation((short) 90);

        Row row0 = sheet.createRow(0);
        Row row1 = sheet.createRow(1);
        Row row2 = sheet.createRow(2);

        row0.setHeight((short) (3 * 255));

        sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 0));
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 1, 2));
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 3, 6));
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 7, 8));
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 9, 10));
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 11, 12));
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 13, 16));
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 17, 19));

        initCell(row0.createCell(1), "Объект", cellHorStyle);
        initCell(row0.createCell(3), "Информация об использованных\nматериалах", cellHorStyle);
        initCell(row0.createCell(7), "Документ М-29", cellHorStyle);
        initCell(row0.createCell(9), "Документ о\nвыполнении\nСМР (КС-2)", cellHorStyle);
        initCell(row0.createCell(11), "Документ о\nвыполнении\nСМР (КС-3)", cellHorStyle);
        initCell(row0.createCell(13), "Подрядчик", cellHorStyle);
        initCell(row0.createCell(17), "Договор", cellHorStyle);

        int cellWightOne = 5;
        int cellWightTwo = 7;
        int cellWightThree = 9;

        initCellWidth(sheet, cellWightOne, row0.createCell(0), "№\nп/п", cellHorStyle);
        initCellWidth(sheet, cellWightOne, row1.createCell(1), "Код", cellVerStyle);
        initCellWidth(sheet, cellWightTwo, row1.createCell(2), "Наименование", cellVerStyle);
        initCellWidth(sheet, cellWightTwo, row1.createCell(3), "Наименование", cellVerStyle);
        initCellWidth(sheet, cellWightThree, row1.createCell(4), "Номенклатурный\nномер", cellVerStyle);
        initCellWidth(sheet, cellWightOne, row1.createCell(5), "Ед. изм.", cellVerStyle);
        initCellWidth(sheet, cellWightOne, row1.createCell(6), "Кол-во", cellVerStyle);
        initCellWidth(sheet, cellWightOne, row1.createCell(7), "Номер", cellVerStyle);
        initCellWidth(sheet, cellWightTwo, row1.createCell(8), "Дата", cellVerStyle);
        initCellWidth(sheet, cellWightOne, row1.createCell(9), "Номер", cellVerStyle);
        initCellWidth(sheet, cellWightTwo, row1.createCell(10), "Дата", cellVerStyle);
        initCellWidth(sheet, cellWightTwo, row1.createCell(11), "Номер", cellVerStyle);
        initCellWidth(sheet, cellWightTwo, row1.createCell(12), "Дата", cellVerStyle);
        initCellWidth(sheet, cellWightOne, row1.createCell(13), "Код", cellVerStyle);
        initCellWidth(sheet, cellWightOne, row1.createCell(14), "ИНН", cellVerStyle);
        initCellWidth(sheet, cellWightOne, row1.createCell(15), "КПП", cellVerStyle);
        initCellWidth(sheet, cellWightThree, row1.createCell(16), "Наименование", cellVerStyle);
        initCellWidth(sheet, cellWightTwo, row1.createCell(17), "Вид", cellVerStyle);
        initCellWidth(sheet, cellWightTwo, row1.createCell(18), "Номер", cellVerStyle);
        initCellWidth(sheet, cellWightOne, row1.createCell(19), "Дата", cellVerStyle);

        for (int i = 1; i < 21; i++) {
            initCell(row2.createCell(i - 1), String.valueOf(i), cellHorStyle);
        }
    }

    // заполнение таблицы данными
    private void filling(Workbook book, ZGAgentExpenseDTO dto) {
        Sheet sheet = book.getSheet("3Г Отчет агента");

        CellStyle styleCell = cellStyle(book, font(book, false, 8));

        Row row3 = sheet.createRow(3);

        List<String> total = dto.getTotal();

        for (int i = 0; i < total.size(); i++) {
            initCell(row3.createCell(i), total.get(i), styleCell);
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