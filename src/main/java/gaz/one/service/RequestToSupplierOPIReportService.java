package gaz.one.service;

import gaz.one.dto.RequestToSupplierOPIDTO;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Iterator;

public class RequestToSupplierOPIReportService {

    public static void main(String[] args) throws IOException {
        RequestToSupplierOPIReportService requestOPI = new RequestToSupplierOPIReportService();
        requestOPI.updateDocument();
    }

    public  void updateDocument() throws IOException {
        // путь к шаблону
        String templatePage = "src/main/resources/templateWord/1.docx";
        // путь и имя сохраняемого файла.
        String outBasic = "1_new.docx";
        // обращение к базе
        RequestToSupplierOPIDataService one = new RequestToSupplierOPIDataService();
        RequestToSupplierOPIDTO reqDTO = one.create();
        // создание(чтение) документа
        XWPFDocument doc = new XWPFDocument(Files.newInputStream(Path.of(templatePage)));
        Iterator<XWPFParagraph> para = doc.getParagraphsIterator();
        Iterator<XWPFTable> table = doc.getTablesIterator();
        // наполнение документа
        paragraph(para, reqDTO);
        table(table, reqDTO);
        // сохранение документа
        FileOutputStream out = new FileOutputStream(outBasic);
        doc.write(out);
        doc.close();

        System.out.println("\nМожно пробовать");
    }
    // поиск и изменение текста в параграфе
    private void paragraph(Iterator<XWPFParagraph> para, RequestToSupplierOPIDTO req) {
        for (XWPFParagraph xwpfParagraph : para.next().getBody().getParagraphs()) {
            for (XWPFRun xwpfRun : xwpfParagraph.getRuns()) {
                String docText = xwpfRun.getText(0);
                docText = docText.replace("инвестпроект", req.getInvestmentProject());
                docText = docText.replace("Наименование ОКС", req.getNameOKS());
                docText = docText.replace("Подрядчик СМР", req.getContractorSMR());
                docText = docText.replace("Поставщик ОПИ", req.getSupplierOPI());
                docText = docText.replace("Отчетный период поставки (год, месяц)", req.getDatePeriod());
                docText = docText.replace("Наименование карьера", req.getNameQuarry());
                docText = docText.replace("ДД.ММ.ГГГГ", req.getDateRequest());
                xwpfRun.setText(docText, 0);
            }
        }
    }
    // добавление данных в таблицу
    private void table(Iterator<XWPFTable> table, RequestToSupplierOPIDTO req) {
        for (XWPFTable tab : table.next().getBody().getTables()) {
            XWPFTableRow row1 = tab.createRow();
            cellInRow(row1, 0, "3");
            cellInRow(row1, 1, "" + req.getViewOPI());
            cellInRow(row1, 2, "" + req.getNameQuarry());
            cellInRow(row1, 3, "" + req.getNumberContractOPI() + " от " + req.getDateContractOPI());
            cellInRow(row1, 4, "" + req.getDeliveryObject());
            cellInRow(row1, 5, "" + req.getContractorSMR());
            cellInRow(row1, 6, "" + req.getCapacity());
        }
    }

    // метод добавления текста в ячейку
    private static void cellInRow(XWPFTableRow tableRow, int cell, String text) {
        XWPFRun runText = styleText(tableRow, cell);
        runText.setText(text);
        tableRow.getCell(cell).removeParagraph(0);
    }

    // стиль текста
    private static XWPFRun styleText(XWPFTableRow tableRow, int cell) {
        XWPFTableCell tCell = tableRow.getCell(cell);
        tCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.TOP);

        XWPFParagraph paragraph = tCell.addParagraph();
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        paragraph.setIndentationFirstLine(0);
        paragraph.setSpacingBefore(0);
        paragraph.setSpacingAfter(0);

        XWPFRun runText = paragraph.createRun();
        runText.setFontFamily("Times New Roman");
        runText.setFontSize(11);

        return runText;
    }
}