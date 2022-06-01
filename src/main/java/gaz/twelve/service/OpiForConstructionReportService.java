package gaz.twelve.service;

import gaz.twelve.dto.OpiForConstructionDTO;
import org.apache.poi.hssf.usermodel.HSSFHeader;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PropertyTemplate;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class OpiForConstructionReportService {
    public static void main(String[] args) throws IOException {

        OpiForConstructionDTO dto = new OpiForConstructionDTO();

        OpiForConstructionDataService data = new OpiForConstructionDataService();
        data.getOpiForConstructionDTO(dto);

        OpiForConstructionReportService tw = new OpiForConstructionReportService();
        tw.createReport("12.xls", dto);
    }

    public void createReport(String file, OpiForConstructionDTO dto) throws IOException {
// создание документа
        Workbook book = new HSSFWorkbook();
// создаём лист в документе
        Sheet sheet = book.createSheet("Справка о движении ОПИ");
// верхний колонтитул
// TODO -> возможно это лишняя запись
        sheet.getHeader().setCenter(HSSFHeader.font("Arial", "bold") + HSSFHeader.fontSize((short) 12)
                + "СПРАВКА О ДВИЖЕНИИ ОПИ, ПРЕДНАЗНАЧЕННЫХ ДЛЯ СТРОИТЕЛЬСТВА ");
// задаем отступ от края листа для печати
        sheet.setMargin(Sheet.LeftMargin, 0.4);
        sheet.setMargin(Sheet.RightMargin, 0.4);
// устанавливаем ориентацию листа для печати (альбомная)
        sheet.getPrintSetup().setLandscape(true);
// выравнивание по центру листа
        sheet.setHorizontallyCenter(true);
// перенос рядов на каждый лист
        sheet.setRepeatingRows(CellRangeAddress.valueOf("19"));

// наполнение документа
        header(book, sheet, dto);
        headerTwo(book, sheet, dto);
        nameColumn(book, sheet);
        filling(book, sheet, dto);

// отображение границ таблицы
        PropertyTemplate propertyTemplate = new PropertyTemplate();
        propertyTemplate.drawBorders(new CellRangeAddress(0, 10, 0, 15), BorderStyle.THIN, BorderExtent.OUTSIDE);
        propertyTemplate.drawBorders(new CellRangeAddress(0, 10, 0, 15), BorderStyle.THIN, BorderExtent.INSIDE_HORIZONTAL);
        propertyTemplate.drawBorders(new CellRangeAddress(12, 14, 11, 16), BorderStyle.THIN, BorderExtent.ALL);
        propertyTemplate.drawBorders(new CellRangeAddress(16, sheet.getLastRowNum(), 0, 17), BorderStyle.THIN, BorderExtent.ALL);
        propertyTemplate.applyBorders(sheet);

// Сохранение документа
        FileOutputStream out = new FileOutputStream(file);
        book.write(out);
        out.close();

        System.out.println("\nМожно пробовать\n12.xls");
    }



    // создание табличного заголовка
    private void header(Workbook book, Sheet sheet, OpiForConstructionDTO dto) {
// стиль для ячейки
        CellStyle cellStyle = cellStyle(book, font(book, false, 9));
        CellStyle cellStyleLeft = cellStyle(book, font(book, false, 9));
        cellStyleLeft.setAlignment(HorizontalAlignment.LEFT);

        rowHeaderName(sheet, 0, "Заказчик:", cellStyleLeft);
        rowHeaderName(sheet, 1, "", cellStyle);
        rowHeaderName(sheet, 2, "Генподрядчик:", cellStyleLeft);
        rowHeaderName(sheet, 3, "", cellStyle);
        rowHeaderName(sheet, 4, "Инвестиционный проект\n(стройка) ", cellStyleLeft);
        rowHeaderName(sheet, 5, "", cellStyle);
        rowHeaderName(sheet, 6, "Код стройки ", cellStyleLeft);
        rowHeaderName(sheet, 7, "Объект ", cellStyleLeft);
        rowHeaderName(sheet, 8, "", cellStyle);
        rowHeaderName(sheet, 9, "Код Объекта", cellStyleLeft);
        rowHeaderName(sheet, 10, "Договор генподряда", cellStyleLeft);

        rowHeaderData(sheet, 0, dto.getCustomer(), cellStyleLeft);
        rowHeaderData(sheet, 2, dto.getContractor(), cellStyleLeft);
        rowHeaderData(sheet, 4, dto.getInvestProject(), cellStyleLeft);
        rowHeaderData(sheet, 6, dto.getCodeConstr(), cellStyleLeft);
        rowHeaderData(sheet, 7, dto.getNameObjConstr(), cellStyleLeft);
        rowHeaderData(sheet, 9, dto.getCodeObj(), cellStyleLeft);
        rowHeaderData(sheet, 10, dto.getGeneralContract(), cellStyleLeft);
        rowHeaderData(sheet, 1, "(наименование, адрес, телефон, факс)", cellStyle);
        rowHeaderData(sheet, 3, "(наименование, адрес, телефон, факс)", cellStyle);
        rowHeaderData(sheet, 5, "(наименование)", cellStyle);
        rowHeaderData(sheet, 8, "(наименование)", cellStyle);
    }

    // создание 2 табличного заголовка
    private void headerTwo(Workbook book, Sheet sheet, OpiForConstructionDTO dto) {
// стиль для ячейки
        CellStyle cellStyle = cellStyle(book, font(book, false, 9));
        CellStyle cellStyleBold = cellStyle(book, font(book, true, 10));

        Row row11 = sheet.createRow(11);
        Row row12 = sheet.createRow(12);
        Row row13 = sheet.createRow(13);
        Row row14 = sheet.createRow(14);
        Row row15 = sheet.createRow(15);

        row11.setHeight((short) (2 * 255));
        row12.setHeight((short) (2 * 255));
        row15.setHeight((short) (2 * 255));

        sheet.addMergedRegion(new CellRangeAddress(12, 12, 0, 9));
        sheet.addMergedRegion(new CellRangeAddress(12, 12, 11, 12));
        sheet.addMergedRegion(new CellRangeAddress(13, 14, 11, 12));
        sheet.addMergedRegion(new CellRangeAddress(12, 12, 13, 14));
        sheet.addMergedRegion(new CellRangeAddress(13, 14, 13, 14));
        sheet.addMergedRegion(new CellRangeAddress(12, 12, 15, 16));
        sheet.addMergedRegion(new CellRangeAddress(13, 14, 0, 9));

        initCell(row12.createCell(0), "СПРАВКА\n" +
                "о движении ОПИ, предназначенных для строительства", cellStyleBold);
        initCell(row12.createCell(11), "Номер документа", cellStyle);
        initCell(row12.createCell(13), "Дата составления", cellStyle);
        initCell(row12.createCell(15), "Отчетный период", cellStyle);
        initCell(row13.createCell(11), dto.getNumDoc(), cellStyle);
        initCell(row13.createCell(13), dto.getDatePreparation(), cellStyle);
        initCell(row13.createCell(15), "с", cellStyle);
        initCell(row13.createCell(16), "по", cellStyle);
        initCell(row13.createCell(0), "к  Отчету о расходе " +
                "основных материалов в строительстве в сопоставлении с " +
                "расходом, определенным по производственным нормам (форма № М-29) ", cellStyle);
        initCell(row14.createCell(15), dto.getPeriodFrom(), cellStyle);
        initCell(row14.createCell(16), dto.getPeriodТо(), cellStyle);
    }
    // наименование колонок основной таблицы
    private void nameColumn(Workbook book, Sheet sheet) {
        CellStyle cellStyle = cellStyle(book, font(book, false, 9));
        CellStyle verStyle = cellStyle(book, font(book, false, 9));
        verStyle.setRotation((short) 90);

        int columWight = 5;
        int columWight2 = 8;

        Row row16 = sheet.createRow(16);
        Row row17 = sheet.createRow(17);
        Row row18 = sheet.createRow(18);


        row16.setHeight((short) (3 * 255));

        sheet.addMergedRegion(new CellRangeAddress(16, 16, 3, 4));
        sheet.addMergedRegion(new CellRangeAddress(16, 16, 15, 16));

        for (int i = 0; i < 18; i++) {
            if (i != 3 && i != 4 && i != 15 && i != 16) {
                sheet.addMergedRegion(new CellRangeAddress(16, 17, i, i));
            }
        }

        initCell(row16.createCell(3), "Единица измерения", cellStyle);
        initCell(row16.createCell(15), "СПРАВОЧНО:\nобъем ОПИ исходя из проектного объема работ", cellStyle);

        initCellWidth(sheet, columWight + 1, row16.createCell(0), "№\nп/п", cellStyle);
        initCellWidth(sheet, columWight, row16.createCell(1), "Наименование материала\n(ОПИ)", verStyle);
        initCellWidth(sheet, columWight, row16.createCell(2), "Номенклатурный номер", verStyle);
        initCellWidth(sheet, columWight, row17.createCell(3), "код", verStyle);
        initCellWidth(sheet, columWight, row17.createCell(4), "наименование", verStyle);
        initCellWidth(sheet, columWight2, row16.createCell(5), "Получено ОПИ\n(в основном состоянии)\nс начала строительства", verStyle);
        initCellWidth(sheet, columWight, row16.createCell(6), "№ накладной", verStyle);
        initCellWidth(sheet, columWight, row16.createCell(7), "Дата накладной", verStyle);
        initCellWidth(sheet, columWight2, row16.createCell(8), "Израсходовано ОПИ\n(в основном состоянии)\nс начала строительства", verStyle);
        initCellWidth(sheet, columWight, row16.createCell(9), "№ КС-2", verStyle);
        initCellWidth(sheet, columWight, row16.createCell(10), "Дата КС-2", verStyle);
        initCellWidth(sheet, columWight2, row16.createCell(11), "Остаток ОПИ\n(в основном состоянии)\nс начала строительства", verStyle);
        initCellWidth(sheet, columWight2, row16.createCell(12), "Остаток ОПИ\n(в основном состоянии)\nна начало отчетного периода", verStyle);
        initCellWidth(sheet, columWight2, row16.createCell(13), "Поступило ОПИ\n(в основном состоянии)\nза отчетный период", verStyle);
        initCellWidth(sheet, columWight2, row16.createCell(14), "Израсходовано ОПИ\n(в основном состоянии)\nза отчетный период", verStyle);
        initCellWidth(sheet, columWight2 + 4, row17.createCell(15), "израсходовано ОПИ\nисходя из проектного\nобъема работ за\nотчетный период", verStyle);
        initCellWidth(sheet, columWight2 + 4, row17.createCell(16), "коэффициент (соотношение)\nобъема ОПИ в\nосновном состоянии\nи ОПИ исходя из", verStyle);
        initCellWidth(sheet, columWight2, row16.createCell(17), "Остаток ОПИ (в основном состоянии)\nна конец отчетного периода", verStyle);
        // нумеруем колонки под наименованием
        for (int i = 1; i < 19; i++) {
            initCell(row18.createCell(i - 1), String.valueOf(i), cellStyle);
        }
    }

    // добавление данных в основную таблицу
    private void filling(Workbook book, Sheet sheet, OpiForConstructionDTO dto) {
        CellStyle cellStyle = cellStyle(book, font(book, false, 9));

        Row row19 = sheet.createRow(19);
        Row row20 = sheet.createRow(20);

        List<String> dummy = dto.getDummy();
        List<String> total = dto.getTotal();

        for (int i = 0; i < dummy.size(); i++) {
            initCell(row19.createCell(i), dummy.get(i), cellStyle);
        }
        for (int i = 0; i < total.size(); i++) {
            initCell(row20.createCell(i + 1), total.get(i), cellStyle);
        }

        initCell(row20.createCell(0), "итого", cellStyle);

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

    // наименование ряда
    private void rowHeaderName(Sheet sheet, int rowNum, String text, CellStyle cellStyle) {
        Row row = sheet.createRow(rowNum);
        Cell cell = row.createCell(0);
        sheet.addMergedRegion(new CellRangeAddress(rowNum, rowNum, 0, 3));
        cell.setCellValue(text);
        cell.setCellStyle(cellStyle);

        if (rowNum == 4) {
            row.setHeight((short) (2 * 255));
        }
    }

    // добавление данных
    private void rowHeaderData(Sheet sheet, int rowNum, String text, CellStyle cellStyle) {
        Row row = sheet.getRow(rowNum);
        Cell cell = row.createCell(4);
        sheet.addMergedRegion(new CellRangeAddress(rowNum, rowNum, 4, 15));
        cell.setCellValue(text);
        cell.setCellStyle(cellStyle);
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
