package poi.word;

import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ReadDoc {
    public static void main(String[] args) {

        readDocXExtensionDocument();
    }

    private static void readDocXExtensionDocument() {
        String templateWord = "src/main/resources/templateWord/1.docx";
        try {
            XWPFDocument doc = new XWPFDocument(Files.newInputStream(Path.of(templateWord)));
            XWPFWordExtractor extractor = new XWPFWordExtractor(doc);
            extractor.setFetchHyperlinks(true);
            String context = extractor.getText();

            Pattern p = Pattern.compile("<(.+?)>");
            Pattern pp = Pattern.compile("#[A-z|0-9]+#");
            Matcher m = pp.matcher(context);
            while (m.find()) {
//                String link = m.group(0); // часть в квадратных скобках
//                String textString = m.group(1); // текст ссылки без скобок
//                context = context.replaceAll(link, ""); // заказ важно. Ссылка, затем текстовая строка
//                context = context.replaceAll(textString, "");
                context = context.replaceAll("#Date#", "дата");
                context = context.replaceAll("#DatePeriod#", "период");

            }
            System.out.println(context);

        }  catch (IOException e) {
            e.printStackTrace();
        }
    }

}
