package reportGaz;

import org.apache.poi.hssf.usermodel.HSSFHeader;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PropertyTemplate;

import java.io.FileOutputStream;
import java.io.IOException;

public class Fifteen {

    public static void main(String[] args) throws IOException {

        Fifteen tw = new Fifteen();
        tw.createReport("15.xls");

    }

    // формирование таблицы
    public void createReport(String file) throws IOException {
        // создание документа
        Workbook book = new HSSFWorkbook();
        // создаём лист в документе
        Sheet sheet = book.createSheet("Сводка ОПИ для ОКС");

        // верхний колонтитул
        // TODO: возможно это лишняя запись
        sheet.getHeader().setCenter(HSSFHeader.font("Arial", "bold")
                + HSSFHeader.fontSize((short) 12) + "ЕЖЕДНЕВНАЯ СВОДКА О ПОСТАВКЕ ОПИ ДЛЯ ОКС");

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
        dataColumn(book, sheet, fontStyle(book));

        // отображение границ таблицы
        PropertyTemplate propertyTemplate = new PropertyTemplate();
        propertyTemplate.drawBorders(new CellRangeAddress(0, 2, 0, 13),BorderStyle.THIN, BorderExtent.ALL);
        // TODO -> заменить рисование границ на всю область данных
        propertyTemplate.drawBorders(new CellRangeAddress(4, 6, 0, 13),BorderStyle.THIN, BorderExtent.ALL);
        propertyTemplate.applyBorders(sheet);

        // Сохранение документа
        book.write(new FileOutputStream(file));

        System.out.println("\nМожно пробовать\n15.xls");
    }

    // заполнение таблицы наименованием колонок
    private void nameColumn(Workbook book, Sheet sheet, Font font) {
        // стиль ячейки
        CellStyle horStyle = cellStyle(book, font);
        CellStyle verStyle = cellStyle(book, font);
        verStyle.setRotation((short) 90);

        Row row0 = sheet.createRow(0);
        Row row1 = sheet.createRow(1);
        Row row2 = sheet.createRow(2);
        Row row3 = sheet.createRow(3);
        Row row4 = sheet.createRow(4);
        Row row5 = sheet.createRow(5);
        Row row6 = sheet.createRow(6);

        row3.setHeight((short) (2 * 255));
        row4.setHeight((short) (3 * 255));
        row5.setHeight((short) (10 * 228));

        // слияние ячеек
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 13));
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 6));
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 7, 13));
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 0, 6));
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 7, 13));
        sheet.addMergedRegion(new CellRangeAddress(3, 3, 0, 13));

        sheet.addMergedRegion(new CellRangeAddress(4, 4, 0, 1));
        sheet.addMergedRegion(new CellRangeAddress(4, 5, 2, 2));
        sheet.addMergedRegion(new CellRangeAddress(4, 4, 3, 4));
        sheet.addMergedRegion(new CellRangeAddress(4, 4, 5, 8));
        sheet.addMergedRegion(new CellRangeAddress(4, 4, 9, 13));


        // добавление значений в ячейки
        initCell(row0.createCell(0), "Инвестпроект", horStyle);
        initCell(row1.createCell(0), "Наименование", horStyle);
        initCell(row1.createCell(7), "Код (шифр)", horStyle);
        initCell(row2.createCell(0), "", horStyle);                                                          // нужно уточнить что сюда вставлять?
        initCell(row2.createCell(7), "", horStyle);                                                          // нужно уточнить что сюда вставлять?
        initCell(row4.createCell(0), "Субподрядный договор на добычу ОПИ, с видом работ «добыча ОПИ на карьере»", horStyle);

//        initCell(row4.createCell(2), "Контрагент (наименование организации)", horStyle);
//        sheet.setColumnWidth(2, 14 * 256);

        initCell(row4.createCell(3), "Карьер", horStyle);
        initCell(row4.createCell(5), "ОКС", horStyle);
        initCell(row4.createCell(9), "Сводка", horStyle);

        int cellWight = 8;

        initCellWidth(sheet,cellWight+6,row4.createCell(2), "Контрагент (наименование организации)", horStyle);
        initCellWidth(sheet,cellWight+5,row5.createCell(0), "Номер договора", verStyle);
        initCellWidth(sheet,cellWight+5,row5.createCell(1), "Дата договора", verStyle);
        initCellWidth(sheet,cellWight,row5.createCell(3), "Субъект РФ", verStyle);
        initCellWidth(sheet,cellWight,row5.createCell(4), "Наименование\nкарьера", verStyle);
        initCellWidth(sheet,cellWight,row5.createCell(5), "Номер этапа\nстроительства", verStyle);
        initCellWidth(sheet,cellWight,row5.createCell(6), "Наименование этапа\nстроительства", verStyle);
        initCellWidth(sheet,cellWight,row5.createCell(7), "Объект капитального\nстроительства", verStyle);
        initCellWidth(sheet,cellWight,row5.createCell(8), "Код (шифр)", verStyle);
        initCellWidth(sheet,cellWight,row5.createCell(9), "Дата, за которую\nподается сводка", verStyle);
        initCellWidth(sheet,cellWight+4,row5.createCell(10), "Вид работ на ОКС,\nдля которых\nдоставлено ОПИ", verStyle);
        initCellWidth(sheet,cellWight,row5.createCell(11), "ОПИ (Материал)", verStyle);
        initCellWidth(sheet,cellWight,row5.createCell(12), "Объем всего, куб.м", verStyle);
        initCellWidth(sheet,cellWight,row5.createCell(13), "Объем всего, тн", verStyle);

        // нумеруем колонки под наименованием
        for (int i = 1; i < 15; i++) {
            initCell(row6.createCell(i - 1), String.valueOf(i), horStyle);
        }
    }

    // заполнение таблицы данными
    private void dataColumn(Workbook book, Sheet sheet, Font font) {
        // задаем стили ячеек
        DataFormat format = book.createDataFormat();
        font.setBold(false);
        CellStyle verStyle = cellStyle(book, font);
        verStyle.setDataFormat(format.getFormat("dd.MM.yyyy"));
        verStyle.setRotation((short) 90);
    }

    private CellStyle cellStyle(Workbook book, Font font) {
        CellStyle style = book.createCellStyle();
        style.setWrapText(true);                                                                                        // перенос текста
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFont(font);
        return style;
    }

    private Font fontStyle(Workbook book) {
        Font font = book.createFont();
        font.setBold(true);
        font.setFontName("Arial");
        font.setFontHeight((short) (9 * 20));                                                                             //размер шрифта -> 9 = (9 / (1/20))
        return font;
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
