package poi.word;

import org.apache.poi.xwpf.usermodel.IRunBody;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

public class WordReplaceTextInTextBox {

    public static void main(String[] args) throws Exception {

        XWPFDocument document = new XWPFDocument(new FileInputStream("1_new.docx"));

        String someWords = "${name}";
//        String someWords = "TextBox";

        for (XWPFParagraph paragraph : document.getParagraphs()) {
            XmlCursor cursor = paragraph.getCTP().newCursor();
            cursor.selectPath("declare namespace w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' .//*/w:txbxContent/w:p/w:r");

            List<XmlObject> ctrsintxtbx = new ArrayList<>();

            while (cursor.hasNextSelection()) {
                cursor.toNextSelection();
                XmlObject obj = cursor.getObject();
                ctrsintxtbx.add(obj);
            }
            for (XmlObject obj : ctrsintxtbx) {
                CTR ctr = CTR.Factory.parse(obj.xmlText());
                //CTR ctr = CTR.Factory.parse(obj.newInputStream());
                XWPFRun bufferrun = new XWPFRun(ctr, (IRunBody) paragraph);
                String text = bufferrun.getText(0);
                if (text != null && text.contains(someWords)) {
                    text = text.replace(someWords, "replaced");
                    bufferrun.setText(text, 0);
                }
                obj.set(bufferrun.getCTR());
            }
        }

        FileOutputStream out = new FileOutputStream("WordReplaceTextInTextBoxNew.docx");
        document.write(out);
        out.close();
        document.close();
    }

}
