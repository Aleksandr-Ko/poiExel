package gaz.six.service;

import gaz.six.dto.ObtainOPIDTO;
import org.apache.poi.hssf.usermodel.HSSFHeader;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PropertyTemplate;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class ObtainOPIReportServiceXls {

    public static void main(String[] args) throws IOException {

        ObtainOPIDTO dto = new ObtainOPIDTO();

        ObtainOPIDataService data = new ObtainOPIDataService();
        data.getObtainOPIDTO(dto);

        ObtainOPIReportServiceXls report = new ObtainOPIReportServiceXls();
        report.createReport("6.xls", dto);

    }

    // создать отчет
    public void createReport(String file, ObtainOPIDTO dto) throws IOException {
        // создание документа
        Workbook book = new HSSFWorkbook();
        // создаём лист в документе
        Sheet sheet = book.createSheet("Получение и контроль ОПИ");
        // верхний колонтитул
        // TODO -> возможно лишняя запись
        sheet.getHeader().setCenter(HSSFHeader.font("Times New Roman", "") + HSSFHeader.fontSize((short) 12) + "Получение и контроль качества ОПИ");
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
        nameColumn(book, sheet);
        addRow(book, sheet, dto);

        // отображение границ таблицы
        PropertyTemplate propertyTemplate = new PropertyTemplate();
        propertyTemplate.drawBorders(new CellRangeAddress(0, sheet.getLastRowNum(), 0, 10), BorderStyle.THIN, BorderExtent.ALL);
        propertyTemplate.applyBorders(sheet);

        // Сохранение документа
        FileOutputStream out = new FileOutputStream(file);
        book.write(out);
        out.close();

        System.out.println("\nМожно пробовать\n6.xls");
    }

    // наименование колонок основной таблицы
    private void nameColumn(Workbook book, Sheet sheet) {

        CellStyle regularStyle = cellStyle(book, font(book));

        Row row0 = sheet.createRow(0);
        Row row1 = sheet.createRow(1);

        int one = 4;
        int two = 13;
        int three = 18;

        initCellWidth(sheet, one, row0.createCell(0), "№\nп/п", regularStyle);
        initCellWidth(sheet, three, row0.createCell(1), "Наименование приобретенного ОПИ (материала)", regularStyle);
        initCellWidth(sheet, three, row0.createCell(2), "Наименование контрагента", regularStyle);
        initCellWidth(sheet, three, row0.createCell(3), "Номер и дата договора/соглашения", regularStyle);
        initCellWidth(sheet, three, row0.createCell(4), "Наименование ОПИ (материала)", regularStyle);
        initCellWidth(sheet, two, row0.createCell(5), "Объем ОПИ (Материала) ", regularStyle);
        initCellWidth(sheet, three, row0.createCell(6), "Номер и дата первичного учетного документа", regularStyle);
        initCellWidth(sheet, three, row0.createCell(7), "Сумма первичного учетного документа, руб.", regularStyle);
        initCellWidth(sheet, two, row0.createCell(8), "Кроме того, НДС, руб.", regularStyle);
        initCellWidth(sheet, two, row0.createCell(9), "С НДС, руб.", regularStyle);
        initCellWidth(sheet, two, row0.createCell(10), "Дата и номер счета-фактуры", regularStyle);

        // нумеруем колонки под наименованием
        for (int i = 1; i < 12; i++) {
            initCell(row1.createCell(i - 1), String.valueOf(i), regularStyle);
        }
    }

    // вставка строк с данными и итогом
    private void addRow(Workbook book, Sheet sheet, ObtainOPIDTO dto) {
        CellStyle regularStyle = cellStyle(book, font(book));

        for (int i = 0; i < 3; i++) {
            addSetDataOPI(sheet, dto.getDummy(), regularStyle, "ОПИ");
            addSetTotal(sheet, dto.getDummy());
        }
    }


    // метод добавление ряда поставщиков
    private void addSetDataOPI(Sheet sheet, List<String> list, CellStyle cellStyle, String text) {

        Row row = sheet.createRow(sheet.getLastRowNum() + 1);

        sheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum() + 1, 0, 0));
        sheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum() + 1, 1, 1));

        for (int j = 0; j < 2; j++) {
            Row rowData = sheet.createRow(sheet.getLastRowNum() + j);
            for (int i = 0; i < list.size(); i++) {
                initCell(rowData.createCell(i + 2), list.get(i), cellStyle);
            }
            initCell(rowData.createCell(0), "num", cellStyle);
            initCell(rowData.createCell(1), text, cellStyle);
        }

    }

    // метод создание итогового ряда
    private void addSetTotal(Sheet sheet, List<String> list) {

        CellStyle regStyle = cellStyle(sheet.getWorkbook(), font(sheet.getWorkbook()));
        CellStyle leftStyle = cellStyle(sheet.getWorkbook(), font(sheet.getWorkbook()));
        leftStyle.setAlignment(HorizontalAlignment.LEFT);

        Row row = sheet.createRow(sheet.getLastRowNum() + 1);
        row.setHeight((short) (3 * 255));

        initCell(row.createCell(0), "ИТОГО стоимость приобретенного ОПИ (Материала)" +
                " с учетом стоимости услуг по контролю качества", leftStyle);


        sheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), 0, 2));

        for (int i = 3; i < 11; i++) {
            if (i < 7 || i == 10) {
                initCell(row.createCell(i), "X", regStyle);
            } else {
                initCell(row.createCell(i), "", regStyle);
            }
        }
        initCell(row.getCell(7), list.get(1), regStyle);
        initCell(row.getCell(8), list.get(2), regStyle);
        initCell(row.getCell(9), list.get(3), regStyle);

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
    private Font font(Workbook book) {
        Font font = book.createFont();
        font.setFontName("Times New Roman");
        font.setFontHeight((short) (10 * 20));
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
