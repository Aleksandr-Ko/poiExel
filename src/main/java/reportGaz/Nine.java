package reportGaz;

import org.apache.poi.hssf.usermodel.HSSFHeader;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PropertyTemplate;

import java.io.FileOutputStream;
import java.io.IOException;

public class Nine {
    public static void main(String[] args) throws IOException {
        Nine nine = new Nine();
        nine.createReport("9.xls");
    }

    public void createReport(String file) throws IOException {
        // создание документа
        Workbook book = new HSSFWorkbook();
        // создаём лист в документе
        Sheet sheet = book.createSheet("3Г Отчет агента");

        // верхний колонтитул
        // TODO -> возможно это лишняя запись
        sheet.getHeader().setCenter(HSSFHeader.font("Arial", "bold") + HSSFHeader.fontSize((short) 12)
                + " 3В ОТЧЕТА АГЕНТА\n"
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

        // отображение границ таблицы
        PropertyTemplate propertyTemplate = new PropertyTemplate();
        // TODO -> заменить рисование границ на всю область данных
        propertyTemplate.drawBorders(new CellRangeAddress(0, 4, 0, 17), BorderStyle.THIN, BorderExtent.ALL);
        propertyTemplate.applyBorders(sheet);

        // Сохранение документа
        book.write(new FileOutputStream(file));

        System.out.println("\nМожно пробовать\n9.xls");
    }

    private void nameColumn(Workbook book, Sheet sheet) {

        CellStyle cellHorStyle = cellStyle(book, font(book, false, 8));

        Row row0 = sheet.createRow(0);
        Row row1 = sheet.createRow(1);
        Row row2 = sheet.createRow(2);
        Row row3 = sheet.createRow(3);
        Row row4 = sheet.createRow(4);

        row0.setHeight((short) (2 * 255));
        row1.setHeight((short) (3 * 255));

        int cellWightOne = 8;
        int cellWightTwo = 16;
        int cellWightThree = 22;

        sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 0));
        sheet.addMergedRegion(new CellRangeAddress(0, 1, 17, 17));

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 1, 2));
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 3, 6));
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 7, 9));
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 10, 13));
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 14, 16));

        initCell(row0.createCell(1), "Стройка", cellHorStyle);
        initCell(row0.createCell(3), "Подрядчик (место хранения)", cellHorStyle);
        initCell(row0.createCell(7), "Договор", cellHorStyle);
        initCell(row0.createCell(10), "Информация о передаче материалов\nПодрядчику СМР/Заказчику (при возврате)", cellHorStyle);
        initCell(row0.createCell(14), "Первичный документ", cellHorStyle);

        initCellWidth(sheet, cellWightOne, row0.createCell(0), "№\nп/п", cellHorStyle);
        initCellWidth(sheet, cellWightOne, row0.createCell(17), "Счет\n(с/счет)\nбух.\nучета", cellHorStyle);
        initCellWidth(sheet, cellWightOne, row1.createCell(1), "Код", cellHorStyle);
        initCellWidth(sheet, cellWightThree, row1.createCell(2), "Наименование", cellHorStyle);
        initCellWidth(sheet, cellWightOne, row1.createCell(3), "Код", cellHorStyle);
        initCellWidth(sheet, cellWightOne, row1.createCell(4), "ИНН", cellHorStyle);
        initCellWidth(sheet, cellWightOne, row1.createCell(5), "КПП", cellHorStyle);
        initCellWidth(sheet, cellWightThree, row1.createCell(6), "Наименование", cellHorStyle);
        initCellWidth(sheet, cellWightOne, row1.createCell(7), "Вид", cellHorStyle);
        initCellWidth(sheet, cellWightTwo, row1.createCell(8), "Номер", cellHorStyle);
        initCellWidth(sheet, cellWightTwo, row1.createCell(9), "Дата", cellHorStyle);
        initCellWidth(sheet, cellWightThree, row1.createCell(10), "Наименование", cellHorStyle);
        initCellWidth(sheet, cellWightThree, row1.createCell(11), "Номенклатурн.\nномер", cellHorStyle);
        initCellWidth(sheet, cellWightOne, row1.createCell(12), "Ед.\nизм.", cellHorStyle);
        initCellWidth(sheet, cellWightOne, row1.createCell(13), "Кол-\nво", cellHorStyle);
        initCellWidth(sheet, cellWightTwo, row1.createCell(14), "Номер", cellHorStyle);
        initCellWidth(sheet, cellWightTwo, row1.createCell(15), "Дата", cellHorStyle);
        initCellWidth(sheet, cellWightTwo, row1.createCell(16), "Отчетный период", cellHorStyle);

        for (int i = 1; i < 19; i++) {
            initCell(row2.createCell(i - 1), String.valueOf(i), cellHorStyle);
        }
    }

    private Font font(Workbook book, boolean bold, int fontSize) {
        Font font = book.createFont();
        font.setFontName("Arial");
        font.setBold(bold);
        font.setFontHeight((short) (fontSize * 20));
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

    private void initCell(Cell cell, String text, CellStyle cellStyle) {
        cell.setCellValue(text);
        cell.setCellStyle(cellStyle);
    }

    private void initCellWidth(Sheet sheet, int wightColum, Cell cell, String text, CellStyle cellStyle) {
        sheet.setColumnWidth(cell.getColumnIndex(), wightColum * 255);
        cell.setCellValue(text);
        cell.setCellStyle(cellStyle);
    }
}
