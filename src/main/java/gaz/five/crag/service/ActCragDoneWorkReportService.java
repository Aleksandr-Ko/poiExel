package gaz.five.crag.service;

import gaz.RussianMoney;
import gaz.five.crag.dto.ActCragDoneWorkDTO;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.core.io.ClassPathResource;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.List;

public class ActCragDoneWorkReportService {
    public static void main(String[] args) throws IOException {
        ActCragDoneWorkReportService report = new ActCragDoneWorkReportService();
        report.updateDocument();
    }

    public void updateDocument() throws IOException {
        // обращение к базе
        ActCragDoneWorkDataService actData = new ActCragDoneWorkDataService();
        ActCragDoneWorkDTO actDTO = actData.getActCragDoneWorkDTO();
        // путь и имя шаблону
        ClassPathResource f = new ClassPathResource("templateWord/ActCragDoneWork.docx");
        String outBasic = "newActCragDoneWork.docx";
        // чтение документа и запись документа
        try (XWPFDocument doc = new XWPFDocument(f.getInputStream());
             FileOutputStream out = new FileOutputStream(outBasic)) {

            // изменение документа
            paragraph(doc, actDTO);
            updateInCell(doc, actDTO);
            addRow(doc, actDTO);
            addRowSum(doc, actDTO);

            // сохранение документа
            doc.write(out);
        }
    }

    // изменение текста
    private void paragraph(XWPFDocument doc, ActCragDoneWorkDTO actDTO) {
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
                        .replace("firstProxy", actDTO.getFirstProxy())
                        .replace("firstPost", actDTO.getFirstPost())
                        .replace("firstFIO", actDTO.getFirstFIO())
                        .replace("secondProxy", actDTO.getSecondProxy())
                        .replace("secondPost", actDTO.getSecondPost())
                        .replace("secondFIO", actDTO.getSecondFIO())
                        .replace("contractDate", actDTO.getContractDate())
                        .replace("contractNum", actDTO.getContractNum())
                        .replace("periodRequest", actDTO.getPeriodRequest())
                        .replace("costWorkTotal", rusMon(actDTO.getCostWorkTotal()))
                        .replace("costWorkNDS", rusMon(actDTO.getCostWorkNDS()))
                        .replace("costWork", rusMon(actDTO.getCostWork()))
                        .replace("numPaymentOrder", rusMon(actDTO.getNumPaymentOrder()))
                        .replace("datePaymentOrder", rusMon(actDTO.getDatePaymentOrder()))
                        .replace("prepayPercentage", rusMon(actDTO.getPrepayPercentage()))
                        .replace("prepayNDS", rusMon(actDTO.getPrepayNDS()))
                        .replace("prepay", rusMon(actDTO.getPrepay()))
                        .replace("deferPay", rusMon(actDTO.getDeferPay()))
                        .replace("paymentNDS", rusMon(actDTO.getPaymentNDS()))
                        .replace("payment", rusMon(actDTO.getPayment()));

                run.setText(text, 0);
            }
        }
    }

    // изменение текста в ячейке
    private void updateInCell(XWPFDocument doc, ActCragDoneWorkDTO actDTO) {
        SimpleDateFormat dfquotes = new SimpleDateFormat("«dd» MMMM yyyy г.");

        XWPFTableCell city = doc.getTableArray(0).getRow(0).getCell(0);
        XWPFTableCell dateRequest = doc.getTableArray(0).getRow(0).getCell(2);
        XWPFTableCell license = doc.getTableArray(1).getRow(0).getCell(1);

        replaceText(city.getParagraphs(), "city", actDTO.getCity());
        replaceText(dateRequest.getParagraphs(), "dateRequest", dfquotes.format(actDTO.getDateRequest()));
        replaceText(license.getParagraphs(), "license", actDTO.getLicense());
    }

    // добавление ряда
    private void addRow(XWPFDocument doc, ActCragDoneWorkDTO actDTO) {
        XWPFTable table = doc.getTableArray(2);

        XWPFTableRow row = table.createRow();

        addCellInRow(row, 0, "1");
        addCellInRow(row, 1, actDTO.getViewCrag());
        addCellInRow(row, 2, actDTO.getCodeOK());
        addCellInRow(row, 3, actDTO.getCodeOPI());
        addCellInRow(row, 4, actDTO.getCragGOST());
        addCellInRow(row, 5, actDTO.getCapacityCrag());
    }

    // добавление ряда в таблицу с деньгами
    private void addRowSum(XWPFDocument doc, ActCragDoneWorkDTO actDTO) {
        XWPFTable table = doc.getTableArray(3);

        XWPFTableRow row = table.createRow();
        XWPFTableRow row1 = table.createRow();

        addCellInRow(row, 0, "1");
        addCellInRow(row, 1, actDTO.getViewCrag());
        addCellInRow(row, 2, actDTO.getCostWork());
        addCellInRow(row, 3, actDTO.getCostWorkNDS());
        addCellInRow(row, 4, actDTO.getCostWorkTotal());

        addCellInRow(row1, 0, "");
        addCellInRow(row1, 1, "Итого:");
        addCellInRow(row1, 2, actDTO.getCostWork());
        addCellInRow(row1, 3, actDTO.getCostWorkNDS());
        addCellInRow(row1, 4, actDTO.getCostWorkTotal());
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

    // метод для написания суммы прописью
    private static String rusMon(String money) {

        RussianMoney rm = new RussianMoney(money);
        String result = rm.num2str();

        return money + " (" + result + ")";
    }
}
