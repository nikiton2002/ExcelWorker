package ru.nikiton;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;

public class ContractorCellStyles {

    HSSFWorkbook wb;

    //Fonts
//    Font middleFont;
    Font bigFont;
    Font bigFontBold;
//    Font bigFontBoldRed;

    //Cell styles
    CellStyle orderNameStyle;
    CellStyle partNameStyle;
    CellStyle mtpNameStyle;
    CellStyle mtpNameErrorStyle;

    CellStyle orderValStyle;
    CellStyle partValStyle;
    CellStyle mtpValStyle;
    CellStyle itogoValStyle;
    CellStyle itogoHrsValStyle;

    public ContractorCellStyles(HSSFWorkbook wb) {
        this.wb = wb;
/*
        middleFont = wb.createFont();
        middleFont.setFontName("Times New Roman");
        middleFont.setFontHeightInPoints((short) 11);
*/

        bigFont = wb.createFont();
        bigFont.setFontName("Times New Roman");
        bigFont.setFontHeightInPoints((short) 14);

        bigFontBold = wb.createFont();
        bigFontBold.setFontName("Times New Roman");
        bigFontBold.setFontHeightInPoints((short) 14);
        bigFontBold.setBold(true);
/*
        bigFontBoldRed = wb.createFont();
        bigFontBoldRed.setFontName("Times New Roman");
        bigFontBoldRed.setFontHeightInPoints((short) 14);
        bigFontBoldRed.setColor(Font.COLOR_RED);
        bigFontBoldRed.setBold(true);
*/

        orderNameStyle = wb.createCellStyle();
        orderNameStyle.setFont(bigFontBold);
        orderNameStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//        orderNameStyle.setBorderTop(CellStyle.BORDER_MEDIUM);
        orderNameStyle.setBorderRight(CellStyle.BORDER_MEDIUM);
//        orderNameStyle.setBorderBottom(CellStyle.BORDER_THIN);
        orderNameStyle.setBorderLeft(CellStyle.BORDER_THIN);

        partNameStyle = wb.createCellStyle();
        partNameStyle.setFont(bigFont);
        partNameStyle.setAlignment(CellStyle.ALIGN_RIGHT);
        partNameStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//        partNameStyle.setBorderTop(CellStyle.BORDER_THIN);
        partNameStyle.setBorderRight(CellStyle.BORDER_MEDIUM);
//        partNameStyle.setBorderBottom(CellStyle.BORDER_THIN);
        partNameStyle.setBorderLeft(CellStyle.BORDER_THIN);

        mtpNameStyle = wb.createCellStyle();
        mtpNameStyle.setFont(bigFontBold);
        mtpNameStyle.setAlignment(CellStyle.ALIGN_RIGHT);
        mtpNameStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        mtpNameStyle.setBorderTop(CellStyle.BORDER_THIN);
        mtpNameStyle.setBorderRight(CellStyle.BORDER_MEDIUM);
        mtpNameStyle.setBorderBottom(CellStyle.BORDER_THIN);
        mtpNameStyle.setBorderLeft(CellStyle.BORDER_THIN);

        orderValStyle = wb.createCellStyle();
        orderValStyle.setFont(bigFont);
        orderValStyle.setAlignment(CellStyle.ALIGN_CENTER);
        orderValStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        orderValStyle.setBorderTop(CellStyle.BORDER_MEDIUM);
        orderValStyle.setBorderRight(CellStyle.BORDER_MEDIUM);
        orderValStyle.setBorderBottom(CellStyle.BORDER_THIN);
        orderValStyle.setBorderLeft(CellStyle.BORDER_MEDIUM);

        partValStyle = wb.createCellStyle();
        partValStyle.setFont(bigFont);
        partValStyle.setAlignment(CellStyle.ALIGN_CENTER);
        partValStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        partValStyle.setBorderTop(CellStyle.BORDER_THIN);
        partValStyle.setBorderRight(CellStyle.BORDER_MEDIUM);
        partValStyle.setBorderBottom(CellStyle.BORDER_THIN);
        partValStyle.setBorderLeft(CellStyle.BORDER_MEDIUM);

        mtpValStyle = wb.createCellStyle();
        mtpValStyle.setFont(bigFontBold);
        mtpValStyle.setAlignment(CellStyle.ALIGN_CENTER);
        mtpValStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        mtpValStyle.setBorderRight(CellStyle.BORDER_MEDIUM);
        mtpValStyle.setBorderLeft(CellStyle.BORDER_MEDIUM);
        mtpValStyle.setBorderBottom(CellStyle.BORDER_THIN);

        itogoValStyle = wb.createCellStyle();
        itogoValStyle.setFont(bigFontBold);
        DataFormat format = wb.createDataFormat();
        itogoValStyle.setDataFormat(format.getFormat("#,##0.00"));
        itogoValStyle.setAlignment(CellStyle.ALIGN_CENTER);
        itogoValStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        itogoValStyle.setBorderTop(CellStyle.BORDER_MEDIUM);
        itogoValStyle.setBorderRight(CellStyle.BORDER_MEDIUM);
        itogoValStyle.setBorderBottom(CellStyle.BORDER_MEDIUM);
        itogoValStyle.setBorderLeft(CellStyle.BORDER_MEDIUM);

        itogoHrsValStyle = wb.createCellStyle();
        itogoHrsValStyle.setFont(bigFontBold);
        itogoHrsValStyle.setAlignment(CellStyle.ALIGN_CENTER);
        itogoHrsValStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        itogoHrsValStyle.setBorderTop(CellStyle.BORDER_MEDIUM);
        itogoHrsValStyle.setBorderRight(CellStyle.BORDER_MEDIUM);
        itogoHrsValStyle.setBorderBottom(CellStyle.BORDER_MEDIUM);
        itogoHrsValStyle.setBorderLeft(CellStyle.BORDER_MEDIUM);
    }
}
