import org.apache.commons.codec.binary.Hex;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.IOException;

import static org.junit.jupiter.api.Assertions.*;


class XSSFCellToHtmlTest {
    @Test
    void cellToHtmlTest() throws IOException {

        XSSFWorkbook workbook = new XSSFWorkbook(getClass().getResourceAsStream("test.xlsx"));
        XSSFSheet sheet = workbook.getSheet("Sheet1");
        XSSFCell cell = sheet.getRow(0).getCell(0);
        System.out.println(cell.getStringCellValue());
        System.out.println(XSSFCellToHtml.getHtmlString(cell));

        assertEquals("<span style=\"\">Hello </span><span style=\" font-style: italic;\">italic</span><span style=\"\"> </span><span style=\" color: #FF0000;\">red-color</span><span style=\"\"> </span><span style=\" font-weight: bold;\">bold</span><span style=\"\"> </span><span style=\" text-decoration: underline;\">underline</span><span style=\"\"> </span><span style=\" text-decoration: line-through;\">Strike-through</span>", XSSFCellToHtml.getHtmlString(cell));


//        <span style="">Hello </span>
//        <span style=" font-style: italic;">italic</span>
//        <span style=""> </span>
//        <span style=" color: #FF0000;">red-color</span>
//        <span style=""> </span>
//        <span style=" font-weight: bold;">bold</span>
//        <span style=""> </span>
//        <span style=" text-decoration: underline;">underline</span>
//        <span style=""> </span>
//        <span style=" text-decoration: line-through;">Strike-through</span>
    }
}