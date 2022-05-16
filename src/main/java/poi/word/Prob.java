package poi.word;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;

public class Prob {
    public static void main(String[] args)throws IOException {

        Prob prob = new Prob();
        prob.updateDocument(
                "1.docx",
                "1_new.docx",
                "ОАО ресурс Природных материалов");

    }

    private void updateDocument(String input, String output, String replaceText)throws IOException {

        try (XWPFDocument doc = new XWPFDocument(Files.newInputStream(Paths.get(input)))) {

            List<XWPFParagraph> xwpfParagraphList = doc.getParagraphs();
            // Iterate over paragraph list and check for the replaceable text in each paragraph -
            // Перебрать список абзацев и проверить наличие заменяемого текста в каждом абзаце.
            for (XWPFParagraph xwpfParagraph : xwpfParagraphList) {
                for (XWPFRun xwpfRun : xwpfParagraph.getRuns()) {
//                for (XWPFRun xwpfRun : xwpfParagraph.getRuns()) {
                    String docText = xwpfRun.getText(0);
                    //replacement and setting position - замена и настройка положения
                    docText = docText.replace("Поставщик ОПИ", replaceText);
                    xwpfRun.setText(docText, 0);
                }
            }
            // save the docs - сохранить документы
            try (FileOutputStream out = new FileOutputStream(output)) {
                doc.write(out);
            }

        }

    }
}
