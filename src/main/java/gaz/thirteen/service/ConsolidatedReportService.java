package gaz.thirteen.service;

import gaz.thirteen.dto.ConsolidatedDTO;
import org.apache.poi.hssf.usermodel.HSSFHeader;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PropertyTemplate;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class ConsolidatedReportService {
    public static void main(String[] args) throws IOException {

        ConsolidatedDTO dto = new ConsolidatedDTO();

        ConsolidatedDataService data = new ConsolidatedDataService();
        data.getConsolidatedDTO(dto);

        ConsolidatedReportService report = new ConsolidatedReportService();
        report.createReport("13.xls", dto);

    }

    // создать отчет
    public void createReport(String file, ConsolidatedDTO dto) throws IOException {
// создание документа
        Workbook book = new HSSFWorkbook();
// создаём лист в документе
        Sheet sheet = book.createSheet("Сводный отчет");
// верхний колонтитул
        sheet.getHeader().setCenter(HSSFHeader.font("Times New Roman", "") + HSSFHeader.fontSize((short) 12) + "Сводный отчет");
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
        header(book, sheet, dto);
        nameColumn(book, sheet);
        addRow(book, sheet, dto);

// отображение границ таблицы
        PropertyTemplate propertyTemplate = new PropertyTemplate();
        propertyTemplate.drawBorders(new CellRangeAddress(0, 2, 17, 24), BorderStyle.THIN, BorderExtent.ALL);
        propertyTemplate.drawBorders(new CellRangeAddress(4, sheet.getLastRowNum(), 0, 24), BorderStyle.THIN, BorderExtent.ALL);
        propertyTemplate.applyBorders(sheet);
// Сохранение документа
        FileOutputStream out = new FileOutputStream(file);
        book.write(out);
        out.close();

        System.out.println("\nМожно пробовать\n13.xls");
    }

    // создание табличного заголовка
    private void header(Workbook book, Sheet sheet, ConsolidatedDTO dto) {
// стили для ячеек
        CellStyle boldStyle = cellStyle(book, font(book, true, 10));
        CellStyle regularStyle = cellStyle(book, font(book, false, 10));
        CellStyle regularStyleLeft = cellStyle(book, font(book, false, 10));
        regularStyleLeft.setAlignment(HorizontalAlignment.LEFT);

        Row row0 = sheet.createRow(0);
        Row row1 = sheet.createRow(1);
        Row row2 = sheet.createRow(2);
        Row row3 = sheet.createRow(3);

        row3.setHeight((short) (2 * 255));

        sheet.addMergedRegion(new CellRangeAddress(0, 2, 0, 1));
        sheet.addMergedRegion(new CellRangeAddress(0, 2, 2, 10));
        sheet.addMergedRegion(new CellRangeAddress(0, 1, 17, 18));
        sheet.addMergedRegion(new CellRangeAddress(0, 1, 19, 20));
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 21, 24));

        sheet.addMergedRegion(new CellRangeAddress(1, 1, 21, 22));
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 23, 24));

        sheet.addMergedRegion(new CellRangeAddress(2, 2, 17, 18));
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 19, 20));
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 21, 22));
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 23, 24));

        initCell(row0.createCell(0), "Агент:", boldStyle);
        initCell(row0.createCell(2), dto.getAgent(), regularStyleLeft);
        initCell(row0.createCell(17), "Номер\nдокумента", regularStyle);
        initCell(row0.createCell(19), "Дата\nсоставления", regularStyle);
        initCell(row0.createCell(21), "Отчетный период", regularStyle);
        initCell(row1.createCell(21), "c", regularStyle);
        initCell(row1.createCell(23), "по", regularStyle);
        initCell(row2.createCell(17), dto.getNumDoc(), regularStyle);
        initCell(row2.createCell(19), dto.getDatePreparation(), regularStyle);
        initCell(row2.createCell(21), dto.getPeriodFrom(), regularStyle);
        initCell(row2.createCell(23), dto.getPeriodTo(), regularStyle);
    }

    // наименование колонок основной таблицы
    private void nameColumn(Workbook book, Sheet sheet) {
        CellStyle regularStyle = cellStyle(book, font(book, false, 10));
        CellStyle verStyle = cellStyle(book, font(book, false, 10));
        verStyle.setRotation((short) 90);

        int weightCell = 9;

        Row row4 = sheet.createRow(4);
        Row row5 = sheet.createRow(5);
        Row row6 = sheet.createRow(6);

        row4.setHeight((short) (3 * 255));

        sheet.addMergedRegionUnsafe(new CellRangeAddress(4, 4, 1, 2));
        sheet.addMergedRegionUnsafe(new CellRangeAddress(4, 4, 3, 4));
        sheet.addMergedRegionUnsafe(new CellRangeAddress(4, 4, 5, 6));
        sheet.addMergedRegionUnsafe(new CellRangeAddress(4, 4, 7, 8));
        sheet.addMergedRegionUnsafe(new CellRangeAddress(4, 4, 9, 10));
        sheet.addMergedRegionUnsafe(new CellRangeAddress(4, 4, 13, 14));
        sheet.addMergedRegionUnsafe(new CellRangeAddress(4, 4, 19, 21));
        sheet.addMergedRegionUnsafe(new CellRangeAddress(4, 4, 22, 23));

        sheet.addMergedRegionUnsafe(new CellRangeAddress(4, 5, 0, 0));
        sheet.addMergedRegionUnsafe(new CellRangeAddress(4, 5, 11, 11));
        sheet.addMergedRegionUnsafe(new CellRangeAddress(4, 5, 12, 12));
        sheet.addMergedRegionUnsafe(new CellRangeAddress(4, 5, 15, 15));
        sheet.addMergedRegionUnsafe(new CellRangeAddress(4, 5, 16, 16));
        sheet.addMergedRegionUnsafe(new CellRangeAddress(4, 5, 17, 17));
        sheet.addMergedRegionUnsafe(new CellRangeAddress(4, 5, 18, 18));
        sheet.addMergedRegionUnsafe(new CellRangeAddress(4, 5, 24, 24));

        initCell(row4.createCell(1), "Стройка", regularStyle);
        initCell(row4.createCell(3), "Объект", regularStyle);
        initCell(row4.createCell(5), "Контрагент\n(подрядчик СМР)", regularStyle);
        initCell(row4.createCell(7), "Договор\nподряда", regularStyle);
        initCell(row4.createCell(9), "Карьер", regularStyle);
        initCell(row4.createCell(13), "Единица измерения", regularStyle);
        initCell(row4.createCell(19), "Израсходовано ОПИ\n(в основном состоянии)\nза отчетный месяц", regularStyle);
        initCell(row4.createCell(22), "СПРАВОЧНО:\n"
                + "объем использованных ОПИ исходя из проектного объема работ", regularStyle);

        initCellWidth(sheet, 4, row4.createCell(0), "№\nп/п", regularStyle);
        initCellWidth(sheet, weightCell, row5.createCell(1), "Код", verStyle);
        initCellWidth(sheet, weightCell, row5.createCell(2), "Наименование ", verStyle);
        initCellWidth(sheet, weightCell, row5.createCell(3), "Код", verStyle);
        initCellWidth(sheet, weightCell, row5.createCell(4), "Наименование", verStyle);
        initCellWidth(sheet, weightCell, row5.createCell(5), "Код", verStyle);
        initCellWidth(sheet, weightCell, row5.createCell(6), "Наименование", verStyle);
        initCellWidth(sheet, weightCell, row5.createCell(7), "Номер", verStyle);
        initCellWidth(sheet, weightCell, row5.createCell(8), "Дата", verStyle);
        initCellWidth(sheet, weightCell, row5.createCell(9), "Код", verStyle);
        initCellWidth(sheet, weightCell, row5.createCell(10), "Наименование карьера", verStyle);
        initCellWidth(sheet, weightCell, row4.createCell(11), "Наименование материала (ОПИ)", verStyle);
        initCellWidth(sheet, weightCell, row4.createCell(12), "Номенклатурный номер", verStyle);
        initCellWidth(sheet, weightCell, row5.createCell(13), "Код", verStyle);
        initCellWidth(sheet, weightCell, row5.createCell(14), "Наименование", verStyle);
        initCellWidth(sheet, 12, row4.createCell(15),
                "Остаток ОПИ (в основном состоянии*)\nна начало отчетного периода", verStyle);
        initCellWidth(sheet, 12, row4.createCell(16),
                "Получено ОПИ (в основном состоянии)\nза отчетный месяц", verStyle);
        initCellWidth(sheet, weightCell, row4.createCell(17), "№ накладной", verStyle);
        initCellWidth(sheet, weightCell, row4.createCell(18), "Дата накладной", verStyle);
        initCellWidth(sheet, weightCell, row5.createCell(19), "количество ОПИ", verStyle);
        initCellWidth(sheet, weightCell, row5.createCell(20), "№, дата Справки о движении\nОПИ к Отчету ф.№М-29", verStyle);
        initCellWidth(sheet, weightCell, row5.createCell(21), "№, дата акта ф.№КС-2", verStyle);
        initCellWidth(sheet, 16, row5.createCell(22), """
                израсходовано ОПИ исходя
                из проектного объема работ
                за отчетный период
                (в соответствии с ф.№КС-2)""", verStyle);
        initCellWidth(sheet, 16, row5.createCell(23), """
                коэффициент (соотношение)
                объема ОПИ в основном
                состоянии и ОПИ исходя из
                проектного объема работ""", verStyle);
        initCellWidth(sheet, weightCell, row4.createCell(24), "Остаток ОПИ (в основном состоянии)\nна конец отчетного периода", verStyle);
// нумеруем колонки под наименованием
        for (int i = 1; i < 26; i++) {
            initCell(row6.createCell(i - 1), String.valueOf(i), regularStyle);
        }
    }

    // вставка строк с данными и итогом
    private void addRow(Workbook book, Sheet sheet, ConsolidatedDTO dto) {

        CellStyle regularStyle = cellStyle(book, font(book, false, 10));
        CellStyle verStyle = cellStyle(book, font(book, false, 10));
        verStyle.setRotation((short) 90);

        addSetData(sheet, dto.getConstruction(), verStyle);

        addSetDataSupply(sheet, dto.getQuarry(), regularStyle, verStyle, "по сторонним\nкарьерам");
        addSetDataSupply(sheet, dto.getSubject(), regularStyle, verStyle, "по\nсубъектам");
        addSetDataSupply(sheet, dto.getProvider(), regularStyle, verStyle, "по\nпоставщикам");
        addSetDataSupply(sheet, dto.getContractor(), regularStyle, verStyle, "по\nподрядчикам");
        addSetDataSupply(sheet, dto.getObject(), regularStyle, verStyle, "по\nобъектам");

        addSetTotal(sheet, dto.getConstructionT(),regularStyle, verStyle);
        addSetTotal(sheet, dto.getQuarryT(),regularStyle, verStyle);
        addSetTotal(sheet, dto.getSubjectT(),regularStyle, verStyle);
        addSetTotal(sheet, dto.getProviderT(),regularStyle, verStyle);
        addSetTotal(sheet, dto.getContractorT(),regularStyle, verStyle);
        addSetTotal(sheet, dto.getObjectT(),regularStyle, verStyle);
    }


    // метод добавление ряда стройки
    private void addSetData(Sheet sheet, List<String> list, CellStyle cellStyle) {

        Row row = sheet.createRow(sheet.getLastRowNum() + 1);

        for (int i = 0; i < list.size(); i++) {
            initCell(row.createCell(i), list.get(i), cellStyle);
        }
    }

    // метод добавление ряда поставщиков
    private void addSetDataSupply(Sheet sheet, List<String> list, CellStyle cellStyle, CellStyle verStyle, String text) {

        Row row = sheet.createRow(sheet.getLastRowNum() + 1);

        sheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), 0, 2));

        row.setHeight((short) (2 * 255));

        initCell(row.createCell(0), text, cellStyle);

        for (int i = 0; i < list.size(); i++) {
            initCell(row.createCell(i+3), list.get(i), verStyle);
        }
    }

    // метод создание итогового ряда
    private void addSetTotal(Sheet sheet, List<String> list, CellStyle regStyle,CellStyle verStyle) {

        Row row = sheet.createRow(sheet.getLastRowNum() + 1);

        initCell(row.createCell(0), "Итого", regStyle);
        sheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), 0, 14));

        for (int i = 15; i < 26; i++) {
            if (i == 16 | i == 17 | i == 20 | i == 23 | i == 25) {
                initCell(row.createCell(i - 1), "х", regStyle);
            } else {
                initCell(row.createCell(i - 1), "", verStyle);
            }
        }
        initCell(row.getCell(17), list.get(0), verStyle);
        initCell(row.getCell(18), list.get(1), verStyle);
        initCell(row.getCell(20), list.get(2), verStyle);
        initCell(row.getCell(21), list.get(3), verStyle);
        initCell(row.getCell(23), list.get(4), verStyle);
    }

    // метод добавление текста в ячейку
    private void initCell(Cell cell, String text, CellStyle cellStyle) {
        cell.setCellValue(text);
        cell.setCellStyle(cellStyle);
    }

    // метод добавление текста в ячейку с указанием ширины колонки
    private void initCellWidth(Sheet sheet, int wight, Cell cell, String text, CellStyle cellStyle) {
        sheet.setColumnWidth(cell.getColumnIndex(), wight * 256);
        cell.setCellValue(text);
        cell.setCellStyle(cellStyle);
    }

    // стиль шрифта
    private Font font(Workbook book, boolean bold, int fontSize) {
        Font font = book.createFont();
        font.setFontName("Times New Roman");
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
}