package ru.nikiton;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;

public class OperatorCellStyles {

    HSSFWorkbook wb;
    //Fonts
    Font middleFont;
    Font bigFont;
    Font bigFontBold;
    Font bigFontBoldRed;

    //Cell styles
    CellStyle orderNameStyle;
    CellStyle partNameStyle;
    CellStyle mtpNameStyle;
    CellStyle mtpNameErrorStyle;

    //Used in OperatorWorker
    CellStyle orderValStyle;
    CellStyle partValStyle;
    CellStyle mtpValStyle;
    //CellStyle partValZeroStyleGreen;

    CellStyle orderPrcStyle;
    CellStyle partPrcStyle;
    CellStyle mtpPrcStyle;
//    CellStyle partPrcZeroStyle;

    public OperatorCellStyles(HSSFWorkbook wb){
        this.wb = wb;

        //Data formats
        DataFormat format = wb.createDataFormat();
        short noZerosFormat = format.getFormat("#");
        short tripleDigitsMoneyFormat = format.getFormat("#,##0.00");

        middleFont = wb.createFont();
        middleFont.setFontName("Times New Roman");
        middleFont.setFontHeightInPoints((short) 11);

        bigFont = wb.createFont();
        bigFont.setFontName("Times New Roman");
        bigFont.setFontHeightInPoints((short) 14);

        bigFontBold = wb.createFont();
        bigFontBold.setFontName("Times New Roman");
        bigFontBold.setFontHeightInPoints((short) 14);
        bigFontBold.setBold(true);

        bigFontBoldRed = wb.createFont();
        bigFontBoldRed.setFontName("Times New Roman");
        bigFontBoldRed.setFontHeightInPoints((short) 14);
        bigFontBoldRed.setColor(Font.COLOR_RED);
        bigFontBoldRed.setBold(true);

        orderNameStyle = wb.createCellStyle();
        orderNameStyle.setFont(bigFontBold);
        orderNameStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        orderNameStyle.setBorderTop(CellStyle.BORDER_MEDIUM);
        orderNameStyle.setBorderRight(CellStyle.BORDER_MEDIUM);
        orderNameStyle.setBorderBottom(CellStyle.BORDER_MEDIUM);
        orderNameStyle.setBorderLeft(CellStyle.BORDER_THIN);

        partNameStyle = wb.createCellStyle();
        partNameStyle.setFont(bigFont);
        partNameStyle.setAlignment(CellStyle.ALIGN_RIGHT);
        partNameStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        partNameStyle.setBorderTop(CellStyle.BORDER_THIN);
        partNameStyle.setBorderRight(CellStyle.BORDER_MEDIUM);
        partNameStyle.setBorderBottom(CellStyle.BORDER_THIN);
        partNameStyle.setBorderLeft(CellStyle.BORDER_THIN);

        mtpNameStyle = wb.createCellStyle();
        mtpNameStyle.setFont(bigFontBold);
        mtpNameStyle.setAlignment(CellStyle.ALIGN_RIGHT);
        mtpNameStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        mtpNameStyle.setBorderTop(CellStyle.BORDER_THIN);
        mtpNameStyle.setBorderRight(CellStyle.BORDER_MEDIUM);
        mtpNameStyle.setBorderBottom(CellStyle.BORDER_THIN);
        mtpNameStyle.setBorderLeft(CellStyle.BORDER_THIN);

        mtpNameErrorStyle = wb.createCellStyle();
        mtpNameErrorStyle.setFont(bigFontBoldRed);
        mtpNameErrorStyle.setAlignment(CellStyle.ALIGN_RIGHT);
        mtpNameErrorStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        mtpNameErrorStyle.setBorderTop(CellStyle.BORDER_THIN);
        mtpNameErrorStyle.setBorderRight(CellStyle.BORDER_MEDIUM);
        mtpNameErrorStyle.setBorderBottom(CellStyle.BORDER_THIN);
        mtpNameErrorStyle.setBorderLeft(CellStyle.BORDER_THIN);

        orderValStyle = wb.createCellStyle();
        orderValStyle.setFont(bigFont);
        orderValStyle.setAlignment(CellStyle.ALIGN_CENTER);
        orderValStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        orderValStyle.setBorderTop(CellStyle.BORDER_MEDIUM);
        orderValStyle.setBorderRight(CellStyle.BORDER_THIN);
        orderValStyle.setBorderBottom(CellStyle.BORDER_MEDIUM);
        orderValStyle.setBorderLeft(CellStyle.BORDER_MEDIUM);
        orderValStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        orderValStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        orderValStyle.setDataFormat(noZerosFormat);

        partValStyle = wb.createCellStyle();
        partValStyle.setFont(middleFont);
        partValStyle.setAlignment(CellStyle.ALIGN_CENTER);
        partValStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        partValStyle.setBorderTop(CellStyle.BORDER_THIN);
        partValStyle.setBorderRight(CellStyle.BORDER_THIN);
        partValStyle.setBorderBottom(CellStyle.BORDER_THIN);
        partValStyle.setBorderLeft(CellStyle.BORDER_MEDIUM);
        partValStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        partValStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        partValStyle.setDataFormat(noZerosFormat);

        mtpValStyle = wb.createCellStyle();
        mtpValStyle.setFont(bigFont);
        mtpValStyle.setAlignment(CellStyle.ALIGN_CENTER);
        mtpValStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        mtpValStyle.setBorderTop(CellStyle.BORDER_THIN);
        mtpValStyle.setBorderRight(CellStyle.BORDER_THIN);
        mtpValStyle.setBorderBottom(CellStyle.BORDER_THIN);
        mtpValStyle.setBorderLeft(CellStyle.BORDER_MEDIUM);
        mtpValStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        mtpValStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        mtpValStyle.setDataFormat(noZerosFormat);

/*        partValZeroStyleGreen = wb.createCellStyle();
        partValZeroStyleGreen.setFont(middleFont);
        partValZeroStyleGreen.setAlignment(CellStyle.ALIGN_CENTER);
        partValZeroStyleGreen.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        partValZeroStyleGreen.setBorderTop(CellStyle.BORDER_THIN);
        partValZeroStyleGreen.setBorderRight(CellStyle.BORDER_THIN);
        partValZeroStyleGreen.setBorderBottom(CellStyle.BORDER_THIN);
        partValZeroStyleGreen.setBorderLeft(CellStyle.BORDER_MEDIUM);
        partValZeroStyleGreen.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        partValZeroStyleGreen.setFillPattern(CellStyle.SOLID_FOREGROUND);*/

        orderPrcStyle = wb.createCellStyle();
        orderPrcStyle.setFont(bigFont);
        orderPrcStyle.setAlignment(CellStyle.ALIGN_CENTER);
        orderPrcStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        orderPrcStyle.setBorderTop(CellStyle.BORDER_MEDIUM);
        orderPrcStyle.setBorderRight(CellStyle.BORDER_MEDIUM);
        orderPrcStyle.setBorderBottom(CellStyle.BORDER_MEDIUM);
        orderPrcStyle.setBorderLeft(CellStyle.BORDER_THIN);
        orderPrcStyle.setDataFormat(noZerosFormat);

        partPrcStyle = wb.createCellStyle();
        partPrcStyle.setFont(middleFont);
        partPrcStyle.setAlignment(CellStyle.ALIGN_CENTER);
        partPrcStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        partPrcStyle.setBorderTop(CellStyle.BORDER_THIN);
        partPrcStyle.setBorderRight(CellStyle.BORDER_MEDIUM);
        partPrcStyle.setBorderBottom(CellStyle.BORDER_THIN);
        partPrcStyle.setBorderLeft(CellStyle.BORDER_THIN);
        partPrcStyle.setDataFormat(noZerosFormat);

        mtpPrcStyle = wb.createCellStyle();
        mtpPrcStyle.setFont(bigFontBold);
        mtpPrcStyle.setAlignment(CellStyle.ALIGN_CENTER);
        mtpPrcStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        mtpPrcStyle.setBorderTop(CellStyle.BORDER_THIN);
        mtpPrcStyle.setBorderRight(CellStyle.BORDER_MEDIUM);
        mtpPrcStyle.setBorderBottom(CellStyle.BORDER_THIN);
        mtpPrcStyle.setBorderLeft(CellStyle.BORDER_THIN);
        mtpPrcStyle.setDataFormat(noZerosFormat);

/*        partPrcZeroStyle = wb.createCellStyle();
        partPrcZeroStyle.setFont(middleFont);
        partPrcZeroStyle.setAlignment(CellStyle.ALIGN_CENTER);
        partPrcZeroStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        partPrcZeroStyle.setBorderTop(CellStyle.BORDER_THIN);
        partPrcZeroStyle.setBorderRight(CellStyle.BORDER_MEDIUM);
        partPrcZeroStyle.setBorderBottom(CellStyle.BORDER_THIN);
        partPrcZeroStyle.setBorderLeft(CellStyle.BORDER_THIN);*/
    }
}
