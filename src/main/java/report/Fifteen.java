package report;

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
        // задаем отступ от края листа для печати
        sheet.setMargin(Sheet.TopMargin, 0.4);
        sheet.setMargin(Sheet.BottomMargin, 0.4);
        sheet.setMargin(Sheet.LeftMargin, 0.4 );
        sheet.setMargin(Sheet.RightMargin, 0.4 );
        // устанавливаем ориентацию листа для печати (альбомная)
        PrintSetup print = sheet.getPrintSetup();
        print.setLandscape(true);
        // перенос рядов на каждый лист
        sheet.setRepeatingRows(CellRangeAddress.valueOf("6:8"));

        // Наполнение документа
        header(book, sheet, fontStyle(book));
        nameColumn(book, sheet, fontStyle(book));
        dataColumn(book, sheet, fontStyle(book));

        // отображение границ таблицы
        PropertyTemplate propertyTemplate = new PropertyTemplate();
        propertyTemplate.drawBorders(new CellRangeAddress(1, 3, 0, 13),
                BorderStyle.THIN, BorderExtent.ALL);
        // TODO -> заменить рисование границ на всю область данных
        propertyTemplate.drawBorders(new CellRangeAddress(5, 7, 0, 13),
                BorderStyle.THIN, BorderExtent.ALL);
        propertyTemplate.applyBorders(sheet);

        // автовыравнивание высоты ряда для данных (производить один раз перед сохранением!)
        rowDataHeight(sheet);


        // Сохранение документа
        book.write(new FileOutputStream(file));

        System.out.println("\nМожно пробовать");
    }
    // заполнение верхней таблицы
    private void header(Workbook book, Sheet sheet, Font font) {
        // Стиль заголовка
        font.setFontHeight((short)(12*20));
        CellStyle boldStyleBig = cellStyle(book, font);
        Row row0 = sheet.createRow(0);
        row0.setHeight((short)(3*228));
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 13));
        initCell(row0.createCell(0), "ФОРМА ЕЖЕДНЕВНОЙ СВОДКИ О ПОСТАВКЕ ОПИ ДЛЯ ОКС", boldStyleBig);

    }
    // заполнение таблицы наименованием колонок
    private void nameColumn(Workbook book, Sheet sheet, Font font) {
        // стиль ячейки
        CellStyle horStyle = cellStyle(book, font);
        CellStyle verStyle = cellStyle(book, font);
        verStyle.setRotation((short) 90);

        Row row1 = sheet.createRow(1);
        sheet.addMergedRegionUnsafe(new CellRangeAddress(1, 1, 0, 13));
        initCell(row1.createCell(0), "Инвестпроект", horStyle);

        Row row2 = sheet.createRow(2);
        sheet.addMergedRegionUnsafe(new CellRangeAddress(2, 2, 0, 6));
        initCell(row2.createCell(0), "Наименование", horStyle);
        sheet.addMergedRegionUnsafe(new CellRangeAddress(2, 2, 7, 13));
        initCell(row2.createCell(7), "Код (шифр)", horStyle);

        Row row3 = sheet.createRow(3);
        sheet.addMergedRegion(new CellRangeAddress(3, 3, 0, 6));
        initCell(row3.createCell(0), "", horStyle);                                                         // нужно уточнить что сюда вставлять?
        sheet.addMergedRegion(new CellRangeAddress(3, 3, 7, 13));
        initCell(row3.createCell(7), "", horStyle);                                                         // нужно уточнить что сюда вставлять?

        Row row4 = sheet.createRow(4);
        row4.setHeight((short)(2*228));
        sheet.addMergedRegion(new CellRangeAddress(4, 4, 0, 13));

        Row row5 = sheet.createRow(5);
        row5.setHeight((short) (3*228));                                                                                // устанавливаем высоту
        // слияние ячеек
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 0, 1));
        sheet.addMergedRegion(new CellRangeAddress(5, 6, 2, 2));
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 3, 4));
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 5, 8));
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 9, 13));
        // добавление значений в ячейки
        initCell(row5.createCell(0), "Субподрядный договор на добычу ОПИ, с видом работ «добыча ОПИ на карьере»", horStyle);
        initCell(row5.createCell(2), "Контрагент (наименование организации)", horStyle);
        sheet.setColumnWidth(2,14*256);                                                                // устанавливаем ширину
        initCell(row5.createCell(3), "Карьер", horStyle);
        initCell(row5.createCell(5), "ОКС", horStyle);
        initCell(row5.createCell(9), "Сводка", horStyle);

        Row row6 = sheet.createRow(6);
        row6.setHeight((short) (10*228));
        initCell(row6.createCell(0),"Номер договора", verStyle);
        sheet.setColumnWidth(0,13*256);
        initCell(row6.createCell(1),"Дата договора", verStyle);
        sheet.setColumnWidth(1,13*256);
        initCell(row6.createCell(3),"Субъект РФ", verStyle);
        sheet.setColumnWidth(3,8*256);
        initCell(row6.createCell(4),"Наименование\nкарьера", verStyle);
        sheet.setColumnWidth(4,8*256);
        initCell(row6.createCell(5),"Номер этапа\nстроительства", verStyle);
        sheet.setColumnWidth(5,8*256);
        initCell(row6.createCell(6),"Наименование этапа\nстроительства", verStyle);
        sheet.setColumnWidth(6,8*256);
        initCell(row6.createCell(7),"Объект капитального\nстроительства", verStyle);
        sheet.setColumnWidth(7,8*256);
        initCell(row6.createCell(8),"Код (шифр)", verStyle);
        sheet.setColumnWidth(8,8*256);
        initCell(row6.createCell(9),"Дата, за которую\nподается сводка", verStyle);
        sheet.setColumnWidth(9,8*256);
        initCell(row6.createCell(10),"Вид работ на ОКС,\nдля которых\nдоставлено ОПИ", verStyle);
        sheet.setColumnWidth(10,12*256);
        initCell(row6.createCell(11),"ОПИ (Материал)", verStyle);
        sheet.setColumnWidth(11,8*256);
        initCell(row6.createCell(12),"Объем всего, куб.м", verStyle);
        sheet.setColumnWidth(12,8*256);
        initCell(row6.createCell(13),"Объем всего, тн", verStyle);
        sheet.setColumnWidth(13,8*256);
        // нумеруем колонки под наименованием
        Row row7 = sheet.createRow(7);
        for (int i = 1; i < 15; i++) {
            initCell(row7.createCell(i - 1), i, horStyle);
        }
    }
    // заполнение таблицы данными
    private void dataColumn(Workbook book, Sheet sheet, Font font){
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
        font.setFontHeight((short) (9*20));                                                                                //размер шрифта -> 9 = (9 / (1/20))
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

    private void rowDataHeight(Sheet sheet) {
        for(int i = 8; i < sheet.getLastRowNum(); i ++) {
            Row row = sheet.getRow(i);
            int enterCnt = 0;
            int colIdxOfMaxCnt = 1;
            for(int j = 0; j < row.getLastCellNum(); j ++) {
                int rwsTemp = row.getCell(j).toString().split(" ").length;
                if (rwsTemp > enterCnt) {
                    enterCnt = rwsTemp;
                    colIdxOfMaxCnt = j;
                }
            }
            row.setHeight((short)(enterCnt * 228));
        }
    }

}
