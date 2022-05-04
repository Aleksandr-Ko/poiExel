package poi.exel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

public class Begin {
    public static void main(String[] args) throws IOException {
        Begin t = new Begin();

        t.create("Test.xls");

    }

    public void create(String file) throws IOException {


        HSSFWorkbook book = new HSSFWorkbook();
        Sheet sheet = book.createSheet("Birthdays");

        // Нумерация начинается с нуля
        Row row0 = sheet.createRow(0);

        // Мы запишем имя и дату в два столбца
        // имя будет String, а дата рождения --- Date,
        // формата dd.mm.yyyy

        Font font = book.createFont();
        Cell name = row0.createCell(0);

        CellStyle cellStyle = book.createCellStyle();   // создаем стиль ячейки
        cellStyle.setRotation((short) 90);              // задаем стиль ячейки с разворотом
        name.setCellStyle(cellStyle);                   // применяем созданный стиль
        name.setCellValue("John");                      // объявляем текст

        Cell birthdate = row0.createCell(1);

        DataFormat format = book.createDataFormat();
        CellStyle dateStyle = book.createCellStyle();
        dateStyle.setDataFormat(format.getFormat("dd.mm.yyyy"));
        birthdate.setCellStyle(dateStyle);


        // Нумерация лет начинается с 1900-го
        birthdate.setCellValue(new Date(110, 10, 10));

        // Меняем размер столбца
        sheet.autoSizeColumn(1); // Регулирует ширину столбца в соответствии с содержимым (делать в конце обработки)

        // Записываем всё в файл
        book.write(new FileOutputStream(file));
        book.close();
    }
}
