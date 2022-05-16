package gaz.one.service;

import gaz.one.dto.One_1_DTO;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;

public class One_1_ReportService {


    public static void main(String[] args) throws IOException {
        One_1_ReportService one_1ReportService = new One_1_ReportService();
        one_1ReportService.updateDocument("1.docx", "1_new.docx");
    }

    public void updateDocument(String input, String output) throws IOException {

        XWPFDocument doc = new XWPFDocument(Files.newInputStream(Paths.get(input)));


        overwriting(doc);

        //  сохранить документы
        FileOutputStream out = new FileOutputStream(output);
        doc.write(out);

        System.out.println("\nМожно пробовать\n1_new.docx");

    }

    public void overwriting(XWPFDocument doc) {

        One_1_DataService one = new One_1_DataService();
        One_1_DTO one_1_dto = one.create();

        String searchText_IP = "investmentProject";
        String replaceText_IP = one_1_dto.getInvestmentProject();
        String searchText_OKS = "OKS";
        String replaceText_OKS = one_1_dto.getOKS();
        String searchText_SMR = "SMR";
        String replaceText_SMR = one_1_dto.getSMR();
        String searchText_FIO = "FIO";
        String replaceText_FIO = one_1_dto.getFIO();
        String searchText_IO = "IO";
        String replaceText_IO = one_1_dto.getIO();
        String searchText_Date = "Date";
        String replaceText_Date = one_1_dto.getDateRequest();
        String searchText_OPI = "OPI";
        String replaceText_OPI = one_1_dto.getOPI();
        String searchText_Quarry = "Quarry";
        String replaceText_Quarry = one_1_dto.getQuarry();

        List<XWPFParagraph> xwpfParagraphList = doc.getParagraphs();

        for (XWPFParagraph xwpfParagraph : xwpfParagraphList) {
            for (XWPFRun xwpfRun : xwpfParagraph.getRuns()) {

                String docText = xwpfRun.getText(0);

                docText = docText.replace(searchText_IP, replaceText_IP);
                docText = docText.replace(searchText_OKS, replaceText_OKS);
                docText = docText.replace(searchText_SMR, replaceText_SMR);
                docText = docText.replace(searchText_FIO, replaceText_FIO);
                docText = docText.replace(searchText_IO, replaceText_IO);
                docText = docText.replace(searchText_OPI, replaceText_OPI);
                docText = docText.replace(searchText_Date,replaceText_Date);

                docText = docText.replace(searchText_Quarry,replaceText_Quarry);



                xwpfRun.setText(docText, 0);
            }
        }

    }


}