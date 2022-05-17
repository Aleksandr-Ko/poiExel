package gaz.one.service;

import gaz.one.dto.AppendixDTO;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Iterator;
import java.util.List;

public class AppendixReportService {

    public static void main(String[] args) throws IOException, InvalidFormatException {
        AppendixReportService ap = new AppendixReportService();
        ap.updateDocument();

    }

    public void updateDocument() throws IOException {
        // путь к шаблону
        String wayToWordTemplate = "src/main/resources/templateWord/1-1.docx";
        // путь и имя сохраняемого файла.
        String outBasic = "1-1_new.docx";
        // создание(чтение) документа
        XWPFDocument doc = new XWPFDocument(Files.newInputStream(Path.of(wayToWordTemplate)));

        // наполнение документа
        modification(doc);

        //  сохранение документа
        FileOutputStream out = new FileOutputStream(outBasic);
        doc.write(out);

        System.out.println("\nМожно пробовать\n1-1_new.docx");
    }

    private void modification(XWPFDocument doc) {

        AppendixDataService appendixDataService = new AppendixDataService();
        AppendixDTO appendixDTO = appendixDataService.create();

        // создаем ArrayList с помощью iterator (IBodyElement -> XWPFParagraph)
        Iterator<IBodyElement> docElementsIterator = doc.getBodyElementsIterator();

        List<XWPFParagraph> xwpfParagraphList = docElementsIterator.next().getBody().getParagraphs();
        List<XWPFTable> xwpfTableList = docElementsIterator.next().getBody().getTables();

        for (XWPFParagraph xwpfParagraph : xwpfParagraphList) {
            for (XWPFRun xwpfRun : xwpfParagraph.getRuns()) {
                String docText = xwpfRun.getText(0);
                docText = docText.replace("FIO", appendixDTO.getFIO());
                docText = docText.replace("Date", appendixDTO.getDatePeriod());
                xwpfRun.setText(docText, 0);
            }
        }

        for (XWPFTable tab : xwpfTableList) {
            XWPFTableRow row1 = tab.createRow();

            cellInRow(row1, 0, "3");
            cellInRow(row1, 1, "" + appendixDTO.getViewOPI());
            cellInRow(row1, 2, "" + appendixDTO.getNameQuarry());
            cellInRow(row1, 3, "" + appendixDTO.getNumberContractOPI() + " от " + appendixDTO.getDateContractOPI());
            cellInRow(row1, 4, "" + appendixDTO.getDeliveryObject());
            cellInRow(row1, 5, "" + appendixDTO.getContractorSMR());
            cellInRow(row1, 6, "" + appendixDTO.getCapacity());
        }
    }

    private static void cellInRow(XWPFTableRow tableRow, int cell, String text) {
        XWPFRun runText = styleText(tableRow, cell);
        runText.setText(text);
        tableRow.getCell(cell).removeParagraph(0);
    }

    // стиль ячейки в ряду с отступами и выравниванием
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
