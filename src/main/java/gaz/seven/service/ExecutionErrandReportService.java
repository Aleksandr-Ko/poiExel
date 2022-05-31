package gaz.seven.service;

import gaz.seven.dto.ExecutionErrandDTO;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.core.io.ClassPathResource;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.List;

public class ExecutionErrandReportService {

    public static void main(String[] args) throws IOException {
        ExecutionErrandReportService report = new ExecutionErrandReportService();
        report.updateDocument();
    }

    public void updateDocument() throws IOException {
// обращение к сервисам
        ExecutionErrandDataService dataExErr = new ExecutionErrandDataService();
        ExecutionErrandDTO exErrDTO = dataExErr.getExecutionErrandDTO();
// путь и имя шаблону
        ClassPathResource f = new ClassPathResource("templateWord/ExecutionErrand.docx");
        String outBasic = "new_ExecutionErrand.docx";
// чтение документа и запись документа
        try (XWPFDocument doc = new XWPFDocument(f.getInputStream());
             FileOutputStream out = new FileOutputStream(outBasic)) {

// изменение документа
            paragraph(doc, exErrDTO);
            updateInCell(doc, exErrDTO);
            addRow(doc);

// сохранение документа
            doc.write(out);
        }
        System.out.println("\n \\\\ new_ExecutionErrand //");
    }

    // изменение текста
    private void paragraph(XWPFDocument doc, ExecutionErrandDTO dto) {
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
                        .replace("costWorkNDS", dto.getCostWorkNDS())
                        .replace("costWork", dto.getCostWork())
                        .replace("priceLicenseTotal", dto.getPriceLicenseTotal())
                        .replace("priceLicenseNDS", dto.getPriceLicenseNDS())
                        .replace("priceLicense", dto.getPriceLicense())
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
    private void updateInCell(XWPFDocument doc, ExecutionErrandDTO dto) {
        SimpleDateFormat dfQuotes = new SimpleDateFormat("«dd» MMMM yyyy г.");

        XWPFTableCell city = doc.getTableArray(0).getRow(0).getCell(0);
        XWPFTableCell dateRequest = doc.getTableArray(0).getRow(0).getCell(2);

        replaceText(city.getParagraphs(), "city", dto.getCity());
        replaceText(dateRequest.getParagraphs(), "dateRequest", dfQuotes.format(dto.getDateRequest()));
    }

    // добавление ряда
    private void addRow(XWPFDocument doc) {
        XWPFTable table1 = doc.getTableArray(1);
        XWPFTable table2 = doc.getTableArray(2);
        XWPFTable table3 = doc.getTableArray(3);
        XWPFTable table4 = doc.getTableArray(3);

        XWPFTableRow row1 = table1.createRow();
        XWPFTableRow row2 = table2.createRow();
        XWPFTableRow row3 = table3.createRow();
        XWPFTableRow row4 = table4.createRow();


        for (int i = 0; i < row1.getTableCells().size(); i++) {
            addCellInRow(row1, i, "1");
        }
        for (int i = 0; i < row2.getTableCells().size(); i++) {
            addCellInRow(row2, i, "2");
        }
        for (int i = 0; i < row3.getTableCells().size(); i++) {
            addCellInRow(row3, i, "3");
        }
        for (int i = 0; i < row4.getTableCells().size(); i++) {
            if (i == 0) {
                addCellInRow(row4, i, "");
            } else if (i == 1) {
                addCellInRow(row4, i, "Итого:");
            } else if (i == 3) {
                addCellInRow(row4, i, "23649.51");
            }else{
                addCellInRow(row4, i, "X");
            }
        }
    }


    // метод для замены текста в ячейке
    private void replaceText(List<XWPFParagraph> listParagraph, String oldVal, String newVal) {
        for (XWPFParagraph para : listParagraph) {
            for (XWPFRun run : para.getRuns()) {
//                run.setColor("000000");
                run.setFontSize(12);
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