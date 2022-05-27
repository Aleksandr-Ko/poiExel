package gaz.five.sand.service;

import gaz.five.sand.dto.ActSandDoneWorkDTO;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.springframework.core.io.ClassPathResource;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.List;

public class ActSandDoneWorkReportService {

    public static void main(String[] args) throws IOException {
        ActSandDoneWorkReportService report = new ActSandDoneWorkReportService();
        report.updateDocument();
    }


    public void updateDocument() throws IOException {
        // обращение к базе
        ActSandDoneWorkDataService actData = new ActSandDoneWorkDataService();
        ActSandDoneWorkDTO actDTO = actData.getActSandDoneWorkDTO();
        // путь к шаблону
        ClassPathResource f = new ClassPathResource("templateWord/ActSandDoneWork.docx");
        String outBasic = "newActSandDoneWork.docx";
        // чтение документа и запись документа
        try (XWPFDocument doc = new XWPFDocument(f.getInputStream());
             FileOutputStream out = new FileOutputStream(outBasic)) {

            // изменение документа
            paragraph(doc, actDTO);
//            updateInCell(doc, actDTO);

            // сохранение документа
            doc.write(out);
        }
    }

    // изменение текста
    private void paragraph(XWPFDocument doc, ActSandDoneWorkDTO actDTO) {
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
                        .replace("capacitySand", actDTO.getCapacitySand())
                        .replace("costWorkTotal", actDTO.getCostWorkTotal())
                        .replace("costWorkNDS", actDTO.getCostWorkNDS())
                        .replace("costWork", actDTO.getCostWork())
                        .replace("numPaymentOrder", actDTO.getNumPaymentOrder())
                        .replace("datePaymentOrder", actDTO.getDatePaymentOrder())
                        .replace("prepayPercentage", actDTO.getPrepayPercentage())
                        .replace("prepayNDS", actDTO.getPrepayNDS())
                        .replace("prepay", actDTO.getPrepay())
                        .replace("deferPay", actDTO.getDeferPay())
                        .replace("paymentNDS", actDTO.getPaymentNDS())
                        .replace("payment", actDTO.getPayment());

                run.setText(text, 0);
            }
        }
    }
}
