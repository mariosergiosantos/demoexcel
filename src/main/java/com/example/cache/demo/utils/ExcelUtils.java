package com.example.cache.demo.utils;

import static java.math.BigDecimal.ZERO;
import static java.util.Objects.nonNull;
import static org.apache.poi.xssf.usermodel.XSSFFont.DEFAULT_FONT_NAME;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.Iterator;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {

    public static final short TITLE_FONT_SIZE = 16;
    public static final short FONT_SIZE = 11;
    public static final int LAST_COL_TITLE = 5;
    public static final String DEFAULT_EMPTY = "";

    private ExcelUtils() {
    }

    public static XSSFCellStyle defineStyleHeader(XSSFWorkbook workbook) {
        XSSFCellStyle styleHeader = workbook.createCellStyle();
        styleHeader.setFont(defineFontDefault(workbook.createFont()));
        styleHeader.setFillBackgroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        styleHeader.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        styleHeader.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        styleHeader.setAlignment(HorizontalAlignment.CENTER);
        return styleHeader;
    }

    public static XSSFFont defineFontDefault(XSSFFont defaultFont) {
        defaultFont.setFontHeightInPoints(FONT_SIZE);
        defaultFont.setFontName(DEFAULT_FONT_NAME);
        defaultFont.setColor(IndexedColors.BLACK.getIndex());
        defaultFont.setBold(false);
        defaultFont.setItalic(false);
        return defaultFont;
    }

    public static XSSFFont defineFontTitle(XSSFFont fontTitle) {
        fontTitle.setFontHeightInPoints(TITLE_FONT_SIZE);
        fontTitle.setFontName(DEFAULT_FONT_NAME);
        fontTitle.setColor(IndexedColors.BLACK.getIndex());
        fontTitle.setBold(true);
        fontTitle.setItalic(false);
        return fontTitle;
    }

    public static void createHeaderExcel(XSSFSheet firstSheet, String subTitleHeader,
                                         XSSFWorkbook workbook) {
        CellRangeAddress cellRangeAddress = new CellRangeAddress(ZERO.intValue(), ZERO.intValue(), ZERO.intValue(),
                LAST_COL_TITLE);
        firstSheet.addMergedRegion(cellRangeAddress);

        XSSFCellStyle styleTitleHeader = workbook.createCellStyle();
        styleTitleHeader.setFont(defineFontTitle(workbook.createFont()));

        Row row = firstSheet.createRow(ZERO.intValue());
        Cell cellTitle = row.createCell(ZERO.intValue());
        cellTitle.setCellStyle(styleTitleHeader);
        cellTitle.setCellValue(subTitleHeader);

    }

    public static void defineCellInRow(Row row, List<Object> objectValues, XSSFCellStyle style) {
        AtomicInteger colNum = new AtomicInteger();

        objectValues.forEach(field -> {
            Cell cell = row.createCell(colNum.getAndIncrement());
            cell.setCellStyle(style);

            if (field instanceof String) {
                cell.setCellValue((String) field);
            } else if (field instanceof Integer) {
                cell.setCellValue((Integer) field);
            } else if (field instanceof Long) {
                cell.setCellValue((Long) field);
            } else if (field instanceof BigInteger) {
                cell.setCellValue(((BigInteger) field).intValue());
            } else if (field instanceof BigDecimal) {
                cell.setCellValue(((BigDecimal) field).doubleValue());
            } else if (field instanceof Double) {
                cell.setCellValue((Double) field);
            } else {
                cell.setCellValue(DEFAULT_EMPTY);
            }

        });
    }

    public static void defineAutoSizeColumn(XSSFSheet firstSheet, int initRow) {
        if (nonNull(firstSheet) && firstSheet.getPhysicalNumberOfRows() > ZERO.intValue()) {
            Row row = firstSheet.getRow(initRow);
            Iterator<Cell> cellIterator = row.cellIterator();
            cellIterator.forEachRemaining(cell -> firstSheet.autoSizeColumn(cell.getColumnIndex()));
        }
    }
}