package gaz.four.clay.service;

import gaz.four.clay.dto.ActClayReceptionDTO;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.core.io.ClassPathResource;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.List;

public class ActClayReceptionReportService {

    public static void main(String[] args) throws IOException {
        ActClayReceptionReportService report = new ActClayReceptionReportService();
        report.updateDocument();
    }
    // изменение документа
    public void updateDocument() throws IOException {
        // обращение к базе8
        ActClayReceptionDataService actData = new ActClayReceptionDataService();
        ActClayReceptionDTO actDTO = actData.getActClayReceptionDTO();
        // путь к шаблону
        ClassPathResource f = new ClassPathResource("templateWord/ActClayReception.docx");
        // чтение документа
        XWPFDocument doc = new XWPFDocument(f.getInputStream());

        // изменение документа
        paragraph(doc, actDTO);
        updateInCell(doc, actDTO);
        addRow(doc, actDTO);

        // путь и имя сохраняемого файла.
        String outBasic = "newActClayReception.docx";
        // сохранение документа
        FileOutputStream out = new FileOutputStream(outBasic);
        doc.write(out);
        out.close();

    }

    // изменение текста
    private void paragraph(XWPFDocument doc, ActClayReceptionDTO actDTO) {
        // формат даты
        SimpleDateFormat dfMY = new SimpleDateFormat(" MMMM yyyy г.");

        List<XWPFParagraph> listParagraph = doc.getParagraphs();
        for (XWPFParagraph para : listParagraph) {
            for (XWPFRun run : para.getRuns()) {
//                run.setColor("000000");
                String text = run.getText(0);
                text = text.replace("numQuarry", "" + actDTO.getNumQuarry())
                        .replace("nameQuarry", actDTO.getNameQuarry())
                        // TODO ->  какую дату выводить? (запрос, составления или отчетного периода)
                        .replace("dateRequest", dfMY.format(actDTO.getDateRequest()))
                        .replace("firstPost", actDTO.getFirstPost())
                        .replace("firstFIO", actDTO.getFirstFIO())
                        .replace("firstProxy", actDTO.getFirstProxy())
                        .replace("secondPost", actDTO.getSecondPost())
                        .replace("secondFIO", actDTO.getSecondFIO())
                        .replace("secondProxy", actDTO.getSecondProxy())
                        .replace("contractDate", actDTO.getContractDate())
                        .replace("contractNum", actDTO.getContractNum())
                        .replace("periodRequest", actDTO.getPeriodRequest())
                        .replace("capacitySand", actDTO.getCapacityClay())
                        .replace("sheetsNum", actDTO.getSheetsNum());

                run.setText(text, 0);

            }
        }
    }

    // изменение текста в ячейке
    private void updateInCell(XWPFDocument doc, ActClayReceptionDTO actDTO) {
        SimpleDateFormat dfquotes = new SimpleDateFormat("«dd» MMMM yyyy г.");

        XWPFTableCell city = doc.getTableArray(0).getRow(0).getCell(0);
        XWPFTableCell dateRequest = doc.getTableArray(0).getRow(0).getCell(2);
        XWPFTableCell fPost = doc.getTableArray(2).getRow(0).getCell(0);
        XWPFTableCell sPost = doc.getTableArray(2).getRow(0).getCell(1);
        XWPFTableCell fFIO = doc.getTableArray(2).getRow(1).getCell(0);
        XWPFTableCell sFIO = doc.getTableArray(2).getRow(1).getCell(1);

        replaceText(city.getParagraphs(), "city", actDTO.getCity());
        replaceText(dateRequest.getParagraphs(), "dateRequest", dfquotes.format(actDTO.getDateRequest()));
        replaceText(fPost.getParagraphs(), "firstPost", actDTO.getFirstPost());
        replaceText(sPost.getParagraphs(), "secondPost", actDTO.getSecondPost());
        replaceText(fFIO.getParagraphs(), "firstFIO", actDTO.getFirstFIO());
        replaceText(sFIO.getParagraphs(), "secondFIO", actDTO.getSecondFIO());
    }

    // метод добавления ряда
    private static void addCellInRow(XWPFTableRow tableRow, int cell, String text) {
        XWPFRun runText = styleText(tableRow, cell);
        runText.setText(text);
        tableRow.getCell(cell).removeParagraph(0);
    }

    // добавление ряда
    private void addRow(XWPFDocument doc, ActClayReceptionDTO actDTO) {
        XWPFTable table = doc.getTableArray(1);

        XWPFTableRow row = table.createRow();

        addCellInRow(row, 0, "1");
        addCellInRow(row, 1, actDTO.getViewClay());
        addCellInRow(row, 2, actDTO.getCodeOK());
        addCellInRow(row, 3, actDTO.getCodeOPI());
        addCellInRow(row, 4, actDTO.getClayGOST());
        addCellInRow(row, 5, actDTO.getCapacityClay());
    }


    // метод для замены текста в ячейке
    private void replaceText(List<XWPFParagraph> listParagraph, String oldVal, String newVal) {
        for (XWPFParagraph para : listParagraph) {
            for (XWPFRun run : para.getRuns()) {
//                run.setColor("000000");
                run.setFontSize(11);
                String text = run.getText(0);
                text = text.replace(oldVal, newVal);
                run.setText(text, 0);
            }
        }
    }

    // стиль текста
    private static XWPFRun styleText(XWPFTableRow tableRow, int cell) {
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
