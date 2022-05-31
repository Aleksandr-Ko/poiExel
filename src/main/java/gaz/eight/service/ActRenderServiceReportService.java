package gaz.eight.service;

import gaz.eight.dto.ActRenderServiceDTO;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.springframework.core.io.ClassPathResource;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.List;

public class ActRenderServiceReportService {

    public static void main(String[] args)throws IOException {
        ActRenderServiceReportService report = new ActRenderServiceReportService();
        report.updateDocument();
    }
    public void updateDocument() throws IOException {
// обращение к сервисам
        ActRenderServiceDataService dataExErr = new ActRenderServiceDataService();
        ActRenderServiceDTO dto = dataExErr.getActRenderServiceDTO();
// путь и имя шаблону
        ClassPathResource f = new ClassPathResource("templateWord/ActRenderService.docx");
        String outBasic = "new_ActRenderService.docx";
// чтение документа и запись документа
        try (XWPFDocument doc = new XWPFDocument(f.getInputStream());
             FileOutputStream out = new FileOutputStream(outBasic)) {

// изменение документа
            paragraph(doc, dto);
            updateInCell(doc, dto);


// сохранение документа
            doc.write(out);
        }
        System.out.println("\n \\\\ new_ActRenderService //");
    }

    // изменение текста
    private void paragraph(XWPFDocument doc, ActRenderServiceDTO dto) {
// формат даты
        SimpleDateFormat dfMY = new SimpleDateFormat(" MMMM yyyy г.");

        List<XWPFParagraph> listParagraph = doc.getParagraphs();
        for (XWPFParagraph para : listParagraph) {
            for (XWPFRun run : para.getRuns()) {
//                run.setColor("000000");
                String text = run.getText(0);
                text = text.replace("dateRequest", dfMY.format(dto.getDateRequest()))
                        .replace("firstProxy", dto.getFirstProxy())
                        .replace("firstPost", dto.getFirstPost())
                        .replace("firstFIO", dto.getFirstFIO())
                        .replace("secondProxy", dto.getSecondProxy())
                        .replace("secondPost", dto.getSecondPost())
                        .replace("secondFIO", dto.getSecondFIO())
                        .replace("costWorkNDS", dto.getCostWorkNDS())
                        .replace("costWork", dto.getCostWork())
                        .replace("priceLicenseTotal", dto.getPriceLicenseTotal())
                        .replace("priceLicenseNDS", dto.getPriceLicenseNDS())
                        .replace("priceLicense", dto.getPriceLicense())
                        .replace("rewardTotal", dto.getRewardTotal())
                        .replace("rewardNDS", dto.getRewardNDS())
                        .replace("reward", dto.getReward())
// TODO -> какую дату выводить? (запрос, составления или отчетного периода)
                        .replace("contractDate", dto.getContractDate())
                        .replace("contractNum", dto.getContractNum())
                        .replace("periodRequest", dto.getPeriodRequest());

                run.setText(text, 0);
            }
        }
    }

    // изменение текста в ячейке
    private void updateInCell(XWPFDocument doc, ActRenderServiceDTO dto) {
        SimpleDateFormat dfQuotes = new SimpleDateFormat("«dd» MMMM yyyy г.");

        XWPFTableCell city = doc.getTableArray(0).getRow(0).getCell(0);
        XWPFTableCell dateRequest = doc.getTableArray(0).getRow(0).getCell(2);

        replaceText(city.getParagraphs(), "city", dto.getCity());
        replaceText(dateRequest.getParagraphs(), "dateRequest", dfQuotes.format(dto.getDateRequest()));
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

}
