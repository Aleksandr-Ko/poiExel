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

        // ориентация листа для печати (альбомная)
        PrintSetup print = sheet.getPrintSetup();
        print.setLandscape(true);
        // перенос рядов на каждый лист
        sheet.setRepeatingRows(CellRangeAddress.valueOf("6"));

        // Заполнить заголовок
        header(book, sheet, fontStyle(book));
        nameColumn(book, sheet, fontStyle(book));
//            body( book, sheet);

        // отображение границ таблицы
        PropertyTemplate propertyTemplate = new PropertyTemplate();
        propertyTemplate.drawBorders(new CellRangeAddress(1, 3, 0, 13),
                BorderStyle.THIN, BorderExtent.ALL);
        propertyTemplate.drawBorders(new CellRangeAddress(5, 7, 0, 13),
                BorderStyle.THIN, BorderExtent.ALL);
        propertyTemplate.applyBorders(sheet);


        // автовыравнивание ширины колонок
        for (int i = 0; i < 15; i++) {
            sheet.autoSizeColumn(i, true);
        }

        // автовыравнивание высоты колонок
        sheetWeight(book, sheet);


        // Сохранение документа
        book.write(new FileOutputStream(file));

        System.out.println("\nМожно пробовать");
    }

    private void header(Workbook book, Sheet sheet, Font fontBold) {
        // Стиль заголовка
        CellStyle boldStyle = cellStyle(book, fontBold);

        Row row0 = sheet.createRow(0);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 13));
        initCell(row0.createCell(0), "ФОРМА ЕЖЕДНЕВНОЙ СВОДКИ О ПОСТАВКЕ ОПИ ДЛЯ ОКС", boldStyle);

        Row row1 = sheet.createRow(1);
        sheet.addMergedRegionUnsafe(new CellRangeAddress(1, 1, 0, 13));
        initCell(row1.createCell(0), "Инвестпроект", boldStyle);

        Row row2 = sheet.createRow(2);
        sheet.addMergedRegionUnsafe(new CellRangeAddress(2, 2, 0, 6));
        initCell(row2.createCell(0), "Наименование", boldStyle);


        sheet.addMergedRegionUnsafe(new CellRangeAddress(2, 2, 7, 13));
        initCell(row2.createCell(7), "Код (шифр)", boldStyle);


        Row row3 = sheet.createRow(3);
        sheet.addMergedRegion(new CellRangeAddress(3, 3, 0, 6));
        initCell(row3.createCell(0), "", boldStyle);                                                          // нужно уточнить что сюда вставлять?
        sheet.addMergedRegion(new CellRangeAddress(3, 3, 7, 13));
        initCell(row3.createCell(7), "", boldStyle);                                                          // нужно уточнить что сюда вставлять?

        Row row4 = sheet.createRow(4);
        sheet.addMergedRegion(new CellRangeAddress(4, 4, 0, 13));
    }

    private void nameColumn(Workbook book, Sheet sheet, Font font) {

        CellStyle horStyle = cellStyle(book, font);
        CellStyle verStyle = cellStyle(book, font);
        verStyle.setRotation((short) 90);

        Row row6 = sheet.createRow(5);

        sheet.addMergedRegion(new CellRangeAddress(5, 5, 0, 1));
        sheet.addMergedRegion(new CellRangeAddress(5, 6, 2, 2));
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 3, 4));
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 5, 8));
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 9, 13));

        initCell(row6.createCell(0), "Субподрядный договор на\nдобычу ОПИ, с видом работ\n«добыча ОПИ на карьере»", horStyle);
        initCell(row6.createCell(2), "Контрагент\n(наименование\nорганизации)", horStyle);
        initCell(row6.createCell(3), "Карьер", horStyle);
        initCell(row6.createCell(5), "ОКС", horStyle);
        initCell(row6.createCell(9), "Сводка", horStyle);

        Row row7 = sheet.createRow(6);

        initCell(row7.createCell(0),"Номер договора", verStyle);
        initCell(row7.createCell(1),"Дата договора", verStyle);
        initCell(row7.createCell(3),"Субъект РФ", verStyle);
        initCell(row7.createCell(4),"Наименование\nкарьера", verStyle);
        initCell(row7.createCell(5),"Номер этапа\nстроительства", verStyle);
        initCell(row7.createCell(6),"Наименование этапа\nстроительства", verStyle);
        initCell(row7.createCell(7),"Объект капитального\nстроительства", verStyle);
        initCell(row7.createCell(8),"Код (шифр)", verStyle);
        initCell(row7.createCell(9),"Дата, за которую\nподается сводка", verStyle);
        initCell(row7.createCell(10),"Вид работ на ОКС,\nдля которых\nдоставлено ОПИ", verStyle);
        initCell(row7.createCell(11),"ОПИ (Материал)", verStyle);
        initCell(row7.createCell(12),"Объем всего, куб.м", verStyle);
        initCell(row7.createCell(13),"Объем всего, тн", verStyle);


        Row row8 = sheet.createRow(7);
        for (int i = 1; i < 15; i++) {
            initCell(row8.createCell(i - 1), i, horStyle);
        }
    }

    private CellStyle cellStyle(Workbook book, Font font) {
        CellStyle style = book.createCellStyle();
        style.setWrapText(true);                                                                                    // перенос текста
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFont(font);
        return style;
    }

    private Font fontStyle(Workbook book) {
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

    private void sheetWeight(Workbook book,Sheet sheet) {
        for(int i = 1; i < sheet.getLastRowNum(); i ++) {
            Row row = sheet.getRow(i);
            int enterCnt = 0;
            int colIdxOfMaxCnt = 1;
            for(int j = 0; j < row.getLastCellNum(); j ++) {
                int rwsTemp = row.getCell(j).toString().split("\n").length;
                if (rwsTemp > enterCnt) {
                    enterCnt = rwsTemp;
                    colIdxOfMaxCnt = j;
                }
            }
            row.setHeight((short)(enterCnt * 228));
        }

    }

    public final short EXCEL_COLUMN_WIDTH_FACTOR = 256;
    public final int UNIT_OFFSET_LENGTH = 7;
    public final int[] UNIT_OFFSET_MAP = new int[]{0, 36, 73, 109, 146, 182, 219};

    // https://fooobar.com/questions/155344/setting-column-width-in-apache-poi
    public short pixel2WidthUnits(int pxs) {
        short widthUnits = (short) (EXCEL_COLUMN_WIDTH_FACTOR * (pxs / UNIT_OFFSET_LENGTH));
        widthUnits += UNIT_OFFSET_MAP[(pxs % UNIT_OFFSET_LENGTH)];
        return widthUnits;
    }
}
