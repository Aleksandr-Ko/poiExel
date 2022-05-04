package poi.exel.one;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;

public class Test {
    public static void main(String[] args) {

        InputStream is = null;
        Workbook book = null;
        try {
            is = new FileInputStream("C:\\Users\\kopylov\\Desktop\\Project\\work\\gazprom\\gazPrem\\1.xls");
            book = new HSSFWorkbook();

            Sheet sheet = book.getSheetAt(1);
            for (int i = 0; i <= 10; i++) {
                Row row = sheet.getRow(i);
                int enterCnt = 0;
                int colIdxOfMaxCnt = 1;
                for (int j = 0; j <= 10; j++) {
                    int rwsTemp = row.getCell(j).toString().split("\n").length;
                    if (rwsTemp > enterCnt) {
                        enterCnt = rwsTemp;
                        colIdxOfMaxCnt = j;
                    }
                }
                System.out.println(colIdxOfMaxCnt + "Number of rows in the column: " + enterCnt);
                row.setHeight((short) (enterCnt * 228));
            }

            File f = new File("test.xls");
            FileOutputStream out = new FileOutputStream(f);
            book.write(out);

            out.close();
//            is.close();
        } catch (IOException e) {
            return;
        }
    }
}
