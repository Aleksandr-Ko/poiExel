package gaz.two.service;

import gaz.two.dto.ActAcquisitionOPIDTO;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.core.io.ClassPathResource;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.List;

public class ActAcquisitionOPIReportService {
    public static void main(String[] args) throws IOException {
        ActAcquisitionOPIReportService act = new ActAcquisitionOPIReportService();
        act.updateDocument();

    }

    public void updateDocument() throws IOException {
        // обращение к базе
        ActAcquisitionOPIDataService actData = new ActAcquisitionOPIDataService();
        ActAcquisitionOPIDTO actDTO = actData.createDTO();
        // путь к шаблону
        ClassPathResource f = new ClassPathResource("templateWord/ActAcquisitionOPI.docx");
        // чтение документа
        XWPFDocument doc = new XWPFDocument(f.getInputStream());
        List<XWPFParagraph> para = doc.getParagraphs();

        // изменение документа
        paragraph(para, actDTO);
        updateInCell(doc, actDTO);
        addRow(doc, actDTO);
        startForPage_2(doc, 16);

        // путь и имя сохраняемого файла.
        String outBasic = "newActAcquisitionOPI.docx";
        // сохранение документа
        FileOutputStream out = new FileOutputStream(outBasic);
        doc.write(out);
        out.close();
    }


    private void paragraph(List<XWPFParagraph> listParagraph, ActAcquisitionOPIDTO actDTO) {
        // формат даты
        SimpleDateFormat dfMY = new SimpleDateFormat(" MMMM yyyy");

        for (XWPFParagraph para : listParagraph) {
            for (XWPFRun run : para.getRuns()) {
//                run.setColor("000000");
                String text = run.getText(0);
                text = text.replace("firstPost", "" + actDTO.getFirstPost())
                        .replace("firstFIO", "" + actDTO.getFirstFIO())
                        // TODO ->  какую дату выводить? (запрос, составления или отчетного периода)
                        .replace("dateRequest", dfMY.format(actDTO.getDateRequest()))
                        .replace("firstProxy", "" + actDTO.getFirstProxy())
                        .replace("secondPost", "" + actDTO.getSecondPost())
                        .replace("secondFIO", "" + actDTO.getSecondFIO())
                        .replace("secondProxy", "" + actDTO.getSecondProxy())
                        .replace("contractNum", "" + actDTO.getContractNum())
                        .replace("contractDate", "" + actDTO.getContractDate())
                        .replace("periodRequest", "" + actDTO.getPeriodRequest())
                        .replace("expensesNds", "" + actDTO.getExpensesNds())
                        .replace("expenses", "" + actDTO.getExpenses())
                        .replace("rewardTotal", "" + actDTO.getRewardTotal())
                        .replace("rewardNDS", "" + actDTO.getRewardNDS())
                        .replace("reward", "" + actDTO.getReward());

                run.setText(text, 0);
            }
        }
    }

    // добавление данных в таблицу
    private void updateInCell(XWPFDocument doc, ActAcquisitionOPIDTO actDTO) {
        SimpleDateFormat dfquotes = new SimpleDateFormat("«dd» MMMM yyyy г.");

        XWPFTableCell city = doc.getTableArray(0).getRow(0).getCell(0);
        XWPFTableCell dateRequest = doc.getTableArray(0).getRow(0).getCell(2);

        XWPFTableCell fPost = doc.getTableArray(1).getRow(0).getCell(0);
        XWPFTableCell sPost = doc.getTableArray(1).getRow(0).getCell(1);
        XWPFTableCell fFIO = doc.getTableArray(1).getRow(1).getCell(0);
        XWPFTableCell sFIO = doc.getTableArray(1).getRow(1).getCell(1);

        XWPFTableCell fPost1 = doc.getTableArray(4).getRow(0).getCell(0);
        XWPFTableCell sPost1 = doc.getTableArray(4).getRow(0).getCell(1);
        XWPFTableCell fFIO1 = doc.getTableArray(4).getRow(1).getCell(0);
        XWPFTableCell sFIO1 = doc.getTableArray(4).getRow(1).getCell(1);

        replaceText(city.getParagraphs(), "city", actDTO.getCity());
        replaceText(dateRequest.getParagraphs(), "dateRequest", dfquotes.format(actDTO.getDateRequest()));

        replaceText(fPost.getParagraphs(), "firstPost", actDTO.getFirstPost());
        replaceText(sPost.getParagraphs(), "secondPost", actDTO.getSecondPost());
        replaceText(fFIO.getParagraphs(), "firstFIO", actDTO.getFirstFIO());
        replaceText(sFIO.getParagraphs(), "secondFIO", actDTO.getSecondFIO());

        replaceText(fPost1.getParagraphs(), "firstPost", actDTO.getFirstPost());
        replaceText(sPost1.getParagraphs(), "secondPost", actDTO.getSecondPost());
        replaceText(fFIO1.getParagraphs(), "firstFIO", actDTO.getFirstFIO());
        replaceText(sFIO1.getParagraphs(), "secondFIO", actDTO.getSecondFIO());
    }

    private void addRow(XWPFDocument doc, ActAcquisitionOPIDTO actDTO) {
        XWPFTable table = doc.getTableArray(2);
        XWPFTableRow row = table.createRow();
        addCellInRow(row, 0, "1");
        addCellInRow(row, 1, "" + actDTO.getNameOPI());
        addCellInRow(row, 2, "" + actDTO.getPrice());
        addCellInRow(row, 3, "" + actDTO.getPriceNDS());
        addCellInRow(row, 4, "" + actDTO.getPriceTotal());

        XWPFTableRow row1 = table.createRow();
        addCellInRow(row1, 0, "");
        addCellInRow(row1, 1, "Итого за Отчетный период:");
        addCellInRow(row1, 2, "" + actDTO.getPrice());
        addCellInRow(row1, 3, "" + actDTO.getPriceNDS());
        addCellInRow(row1, 4, "" + actDTO.getPriceTotal());

    }


    // метод для начала с новой страницы
    private void startForPage_2(XWPFDocument doc, int paragraph) {
        XWPFParagraph par = doc.getParagraphArray(paragraph);
        par.setPageBreak(true);
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

    // метод добавления ряда
    private static void addCellInRow(XWPFTableRow tableRow, int cell, String text) {
        XWPFRun runText = styleText(tableRow, cell);
        if (cell == 1 ){
            XWPFParagraph par = runText.getParagraph();
            par.setAlignment(ParagraphAlignment.LEFT);
        }
        runText.setText(text);
        tableRow.getCell(cell).removeParagraph(0);
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
