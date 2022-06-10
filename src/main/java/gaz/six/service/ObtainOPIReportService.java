package gaz.six.service;

import gaz.six.dto.ObtainOPIDTO;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.core.io.ClassPathResource;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.List;

public class ObtainOPIReportService {
    public static void main(String[] args) throws IOException {
        ObtainOPIDTO dto = new ObtainOPIDTO();

        ObtainOPIReportService report = new ObtainOPIReportService();
        report.updateDocument(dto);
    }

    public void updateDocument(ObtainOPIDTO dto) throws IOException {
        // обращение к сервисам
        ObtainOPIDataService obtainData = new ObtainOPIDataService();
        ObtainOPIDTO obtainDTO = obtainData.getObtainOPIDTO(dto);
        // путь и имя шаблону
        ClassPathResource f = new ClassPathResource("templateWord/ObtainOPI.docx");
        String outBasic = "new_ObtainOPI.docx";
        // чтение документа и запись документа
        try (XWPFDocument doc = new XWPFDocument(f.getInputStream());
             FileOutputStream out = new FileOutputStream(outBasic)) {

        // изменение документа
            paragraph(doc, obtainDTO);
            updateInCell(doc, obtainDTO);
            addTextCell(doc, obtainDTO);
            addRow(doc);


        // сохранение документа
            doc.write(out);
        }
        System.out.println("\n \\\\ new_ObtainOPI //");
    }

    // изменение текста
    private void paragraph(XWPFDocument doc, ObtainOPIDTO dto) {
        // формат даты
        SimpleDateFormat dfMY = new SimpleDateFormat(" MMMM yyyy г.");

        List<XWPFParagraph> listParagraph = doc.getParagraphs();
        for (XWPFParagraph para : listParagraph) {
            for (XWPFRun run : para.getRuns()) {
//                run.setColor("000000");
                String text = run.getText(0);
                text = text.replace("contractorPost", dto.getContractorPost())
                        .replace("contractorFIO", dto.getContractorFIO())
                        .replace("contractorProxy", dto.getContractorProxy())
                        .replace("costWorkTotal", dto.getCostWorkTotal())
                        .replace("costWorkNDS", dto.getCostWorkNDS())
                        .replace("costWork", dto.getCostWork())
        // TODO -> какую дату выводить? (запрос, составления или отчетного периода)
                        .replace("dateRequest", dfMY.format(dto.getDateRequest()))
                        .replace("contractDate", dto.getContractDate())
                        .replace("contractNum", dto.getContractNum())
                        .replace("periodRequest", dto.getPeriodRequest());

                run.setText(text, 0);
            }
        }
    }

    // изменение текста в ячейке
    private void updateInCell(XWPFDocument doc, ObtainOPIDTO dto) {
        SimpleDateFormat dfQuotes = new SimpleDateFormat("«dd» MMMM yyyy г.");

        XWPFTableCell city = doc.getTableArray(0).getRow(0).getCell(0);
        XWPFTableCell dateRequest = doc.getTableArray(0).getRow(0).getCell(2);

        replaceText(city.getParagraphs(), "city", dto.getCity());
        replaceText(dateRequest.getParagraphs(), "dateRequest", dfQuotes.format(dto.getDateRequest()));
    }

    private void addTextCell(XWPFDocument doc, ObtainOPIDTO dto) {

        XWPFTableRow row2 = doc.getTableArray(2).getRow(2);
        XWPFTableRow row3 = doc.getTableArray(2).getRow(3);
        XWPFTableRow row4 = doc.getTableArray(2).getRow(4);
        XWPFTableRow row5 = doc.getTableArray(2).getRow(5);
        XWPFTableRow row6 = doc.getTableArray(2).getRow(6);
        XWPFTableRow row7 = doc.getTableArray(2).getRow(7);

        addCellInRow(row2, 1, "viewOPI");
        addCellInRow(row5, 1, "viewOPI");

        for (int i = 2; i < row2.getTableCells().size(); i++) {
            addCellInRow(row2, i, "2.1");
            addCellInRow(row3, i, "3.1");
            addCellInRow(row5, i, "5.1");
            addCellInRow(row6, i, "6.1");
            if (i > 4 & i < 8) {
                addCellInRow(row4, i, "4.1");
                addCellInRow(row7, i, "7.1");
            }
        }


    }

    // добавление ряда
    private void addRow(XWPFDocument doc) {
        XWPFTable table1 = doc.getTableArray(1);
        XWPFTable table3 = doc.getTableArray(3);

        XWPFTableRow row1 = table1.createRow();
        XWPFTableRow row3 = table3.createRow();

        for (int i = 0; i < row1.getTableCells().size(); i++) {
            addCellInRow(row1, i, "1");
        }

        for (int i = 0; i < row3.getTableCells().size(); i++) {
            addCellInRow(row3, i, "11");
        }
    }


    // метод для замены текста в ячейке
    private void replaceText(List<XWPFParagraph> listParagraph, String oldVal, String newVal) {
        for (XWPFParagraph para : listParagraph) {
            for (XWPFRun run : para.getRuns()) {
//                run.setColor("000000");
                run.setFontSize(12.5);
                String text = run.getText(0);
                text = text.replace(oldVal, newVal);
                run.setText(text, 0);
            }
        }
    }

    // метод добавления ряда
    private void addCellInRow(XWPFTableRow tableRow, int cell, String text) {
        XWPFRun runText = styleText(tableRow, cell);
        runText.setText(text);
        tableRow.getCell(cell).removeParagraph(0);
    }

    // стиль текста
    private XWPFRun styleText(XWPFTableRow tableRow, int cell) {
        XWPFTableCell tCell = tableRow.getCell(cell);
        tCell.setWidth("auto");
        tCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

        XWPFParagraph paragraph = tCell.addParagraph();
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        paragraph.setIndentationFirstLine(0);
        paragraph.setSpacingBefore(0);
        paragraph.setSpacingAfter(0);

        XWPFRun runText = paragraph.createRun();
        runText.setColor("FF0000");
        runText.setFontFamily("Times New Roman");
        runText.setFontSize(9);

        return runText;
    }
}
