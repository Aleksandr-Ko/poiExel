package gaz.one.service;

import gaz.one.dto.NeedsOPIDTO;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Iterator;

public class NeedsOPIReportService {

    public static void main(String[] args) throws IOException {
        NeedsOPIReportService requestOPI = new NeedsOPIReportService();
        requestOPI.updateDocument();
    }

    public void updateDocument() throws IOException {
        // путь к шаблону
        String templateWord = "src/main/resources/templateWord/NeedsOPI.docx";
        // путь и имя сохраняемого файла.
        String outBasic = "newNeedsOPI.docx";
        // обращение к базе
        NeedsOPIDataService one = new NeedsOPIDataService();
        NeedsOPIDTO reqDTO = one.create();
        // чтение документа
        XWPFDocument doc = new XWPFDocument(Files.newInputStream(Path.of(templateWord)));
        Iterator<XWPFParagraph> para = doc.getParagraphsIterator();
        Iterator<XWPFTable> table = doc.getTablesIterator();
        // изменение документа
        paragraph(para, reqDTO);
        table(table, reqDTO);
        // сохранение документа
        FileOutputStream out = new FileOutputStream(outBasic);
        doc.write(out);
        out.close();

        System.out.println("\nМожно пробовать");
    }

    // поиск и изменение текста в параграфе
    private void paragraph(Iterator<XWPFParagraph> para, NeedsOPIDTO req) {
        for (XWPFParagraph xwpfParagraph : para.next().getBody().getParagraphs()) {
            for (XWPFRun xwpfRun : xwpfParagraph.getRuns()) {
                String docText = xwpfRun.getText(0);

                docText = docText.replace("investmentProject", req.getInvestmentProject())
                        .replace("nameOKS", req.getNameOKS())
                        .replace("contractorSMR", req.getContractorSMR())
                        .replace("supplierOPI", req.getSupplierOPI())
                        .replace("date", req.getDateRequest())
                        .replace("periodDate", req.getDatePeriod())
                        .replace("listNameQuarry", req.getListNameQuarry())
                        .replace("ДД.ММ.ГГГГ", req.getDateRequest());

                xwpfRun.setText(docText, 0);
            }
        }
    }

    // добавление данных в таблицу
    private void table(Iterator<XWPFTable> table, NeedsOPIDTO req) {
        for (XWPFTable tab : table.next().getBody().getTables()) {
            XWPFTableRow row1 = tab.createRow();
            cellInRow(row1, 0, "1");
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
        tCell.setWidth("auto");
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