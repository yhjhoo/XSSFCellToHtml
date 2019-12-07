import org.apache.commons.codec.binary.Hex;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRElt;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRPrElt;

import java.util.Arrays;
import java.util.List;

import static org.apache.commons.collections4.CollectionUtils.isNotEmpty;

public class XSSFCellToHtml {
    public static String getHtmlString(XSSFCell cell) {
        StringBuilder sb = new StringBuilder();
        List<CTRElt> rList = cell.getRichStringCellValue().getCTRst().getRList();
        rList.forEach(ctrElt -> {
            sb.append("<span style=\"");
            formatFont(ctrElt, sb);

            sb.append("\">");
            sb.append(ctrElt.getT());
            sb.append("</span>");
        });

        return sb.toString();
    }

    private static void formatFont(CTRElt ctrElt, StringBuilder sb) {
        CTRPrElt rPr = ctrElt.getRPr();
        if (rPr == null) {
            return;
        }

        if (isNotEmpty(rPr.getBList()) && rPr.getBList().get(0).getVal()) {
            sb.append(" font-weight: bold;");
        }

        if (isNotEmpty(rPr.getIList()) && rPr.getIList().get(0).getVal()) {
            sb.append(" font-style: italic;");
        }

        if (isNotEmpty(rPr.getUList())
                && rPr.getUList().get(0).getVal().toString().equals("single")) {
            sb.append(" text-decoration: underline;");
        }

        if (isNotEmpty(rPr.getStrikeList()) && rPr.getStrikeList().get(0).getVal()) {
            sb.append(" text-decoration: line-through;");
        }

        if (isNotEmpty(rPr.getColorList()) && rPr.getColorList().get(0).xgetRgb() != null) {
            String colorCode = rPr.getColorList().get(0).xgetRgb().getStringValue();

//            if (rPr.getColorList().get(0).getIndexed() > 0) {
//                Long indexed = rPr.getColorList().get(0).getIndexed();
//                byte[] rgbBytes = new DefaultIndexedColorMap().getRGB(indexed.intValue());
//                colorCode = String.valueOf(Hex.encodeHex(rgbBytes));
//            })

            sb.append(" color: #" + colorCode.substring(2) + ";");
        }


    }
}