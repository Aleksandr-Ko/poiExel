package gaz.three.sand.service;

import gaz.three.sand.dto.ActSandShipmentDTO;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.core.io.ClassPathResource;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.List;

public class ActSandShipmentReportService {
    public static void main(String[] args) throws Exception {
        ActSandShipmentReportService report = new ActSandShipmentReportService();
        report.updateDocument();
    }

    public void updateDocument() throws IOException {
        // обращение к базе
        ActSandShipmentDataService actData = new ActSandShipmentDataService();
        ActSandShipmentDTO actDTO = actData.getActSandAcceptanceDTO();
        // путь к шаблону
        ClassPathResource f = new ClassPathResource("templateWord/ActSandShipment.docx");
        // чтение документа
        XWPFDocument doc = new XWPFDocument(f.getInputStream());

        // изменение документа
        paragraph(doc, actDTO);
        updateInCell(doc, actDTO);

        // путь и имя сохраняемого файла.
        String outBasic = "newActSandShipment.docx";
        // сохранение документа
        FileOutputStream out = new FileOutputStream(outBasic);
        doc.write(out);
        out.close();
    }

    // изменение текста
    private void paragraph(XWPFDocument doc, ActSandShipmentDTO actDTO) {
        // формат даты
        SimpleDateFormat dfMY = new SimpleDateFormat(" MMMM yyyy г.");

        List<XWPFParagraph> listParagraph = doc.getParagraphs();
        for (XWPFParagraph para : listParagraph) {
            for (XWPFRun run : para.getRuns()) {
//                run.setColor("000000");
                String text = run.getText(0);
                text = text.replace("numQuarry", "" + actDTO.getNumQuarry())
                        // TODO ->  какую дату выводить? (запрос, составления или отчетного периода)
                        .replace("dateRequest", dfMY.format(actDTO.getDateRequest()))
                        .replace("firstPost", "" + actDTO.getFirstPost())
                        .replace("firstFIO", "" + actDTO.getFirstFIO())
                        .replace("firstProxy", "" + actDTO.getFirstProxy())
                        .replace("secondPost", "" + actDTO.getSecondPost())
                        .replace("secondFIO", "" + actDTO.getSecondFIO())
                        .replace("secondProxy", "" + actDTO.getSecondProxy())
                        .replace("contractDate", "" + actDTO.getContractDate())
                        .replace("contractNum", "" + actDTO.getContractNum())
                        .replace("periodRequest", "" + actDTO.getPeriodRequest())
                        .replace("capacitySand", "" + actDTO.getCapacitySand())
                        .replace("sheetsNum", "" + actDTO.getSheetsNum());

                run.setText(text, 0);
            }
        }
    }

    // изменение текста в ячейке
    private void updateInCell(XWPFDocument doc, ActSandShipmentDTO actDTO) {
        SimpleDateFormat dfquotes = new SimpleDateFormat("«dd» MMMM yyyy г.");

        XWPFTableCell city = doc.getTableArray(0).getRow(0).getCell(0);
        XWPFTableCell dateRequest = doc.getTableArray(0).getRow(0).getCell(2);

        XWPFTableCell nameQuarry = doc.getTableArray(1).getRow(0).getCell(0);

        XWPFTableCell gost = doc.getTableArray(2).getRow(0).getCell(0);
        XWPFTableCell ok = doc.getTableArray(2).getRow(0).getCell(2);
        XWPFTableCell opi = doc.getTableArray(2).getRow(0).getCell(4);

        XWPFTableCell fPost = doc.getTableArray(4).getRow(0).getCell(0);
        XWPFTableCell sPost = doc.getTableArray(4).getRow(0).getCell(1);
        XWPFTableCell fFIO = doc.getTableArray(4).getRow(1).getCell(0);
        XWPFTableCell sFIO = doc.getTableArray(4).getRow(1).getCell(1);

        replaceText(city.getParagraphs(), "city", actDTO.getCity());
        replaceText(dateRequest.getParagraphs(), "dateRequest", dfquotes.format(actDTO.getDateRequest()));
        replaceText(nameQuarry.getParagraphs(), "nameQuarry", actDTO.getNameQuarry());

        replaceText(gost.getParagraphs(), "GOST", actDTO.getStandartGOST());
        replaceText(ok.getParagraphs(), "codeOK", actDTO.getCodeOK());
        replaceText(opi.getParagraphs(), "codeOPI", actDTO.getCodeOPI());

        replaceText(fPost.getParagraphs(), "firstPost", actDTO.getFirstPost());
        replaceText(sPost.getParagraphs(), "secondPost", actDTO.getSecondPost());
        replaceText(fFIO.getParagraphs(), "firstFIO", actDTO.getFirstFIO());
        replaceText(sFIO.getParagraphs(), "secondFIO", actDTO.getSecondFIO());

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
}
