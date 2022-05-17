package gaz.one.service;

import gaz.one.dto.RequestToSupplierOPIDTO;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Iterator;
import java.util.List;

public class RequestToSupplierOPIReportService {

    public static void main(String[] args) throws IOException {
        RequestToSupplierOPIReportService requestOPI = new RequestToSupplierOPIReportService();
        requestOPI.updateDocument();
    }

    public void updateDocument() throws IOException {
        // путь к шаблону
        String wayToWordTemplate = "src/main/resources/templateWord/1.docx";
        // путь и имя сохраняемого файла.
        String outBasic = "1_new.docx";
        // создание(чтение) документа
        XWPFDocument docBasic = new XWPFDocument(Files.newInputStream(Path.of(wayToWordTemplate)));
        // наполнение документа
        modification(docBasic);
        // сохранение документа
        FileOutputStream out = new FileOutputStream(outBasic);
        docBasic.write(out);
        docBasic.close();

        System.out.println("\nМожно пробовать\n1_new.docx");
    }

    private void modification(XWPFDocument doc) {

        RequestToSupplierOPIDataService one = new RequestToSupplierOPIDataService();
        RequestToSupplierOPIDTO requestOPIDTO = one.create();

        // создаем ArrayList с помощью iterator (IBodyElement -> XWPFParagraph)
        Iterator<IBodyElement> docElementsIterator = doc.getBodyElementsIterator();

        List<XWPFParagraph> xwpfParagraphList = docElementsIterator.next().getBody().getParagraphs();

        for (XWPFParagraph xwpfParagraph : xwpfParagraphList) {
            for (XWPFRun xwpfRun : xwpfParagraph.getRuns()) {

                String docText = xwpfRun.getText(0);

                docText = docText.replace("investmentProject", requestOPIDTO.getInvestmentProject());
                docText = docText.replace("OKS", requestOPIDTO.getOKS());
                docText = docText.replace("SMR", requestOPIDTO.getSMR());
                docText = docText.replace("FIO", requestOPIDTO.getFIO());
                docText = docText.replace("IO", requestOPIDTO.getIO());
                docText = docText.replace("OPI", requestOPIDTO.getOPI());
                docText = docText.replace("Date",requestOPIDTO.getDatePeriod());
                docText = docText.replace("Quarry",requestOPIDTO.getQuarry());

                xwpfRun.setText(docText, 0);
            }
        }
    }
}