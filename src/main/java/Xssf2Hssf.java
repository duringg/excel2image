
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.util.HashMap;

public class Xssf2Hssf {

    private int lastColumn = 0;
    private HashMap<Integer, CellStyle> styleMap = new HashMap();

    public void transformXSSF(Workbook workbookOld,
                              Workbook workbookNew) {
        Sheet sheetNew;
        Sheet sheetOld;
        workbookNew.setMissingCellPolicy(workbookOld.getMissingCellPolicy());

        for (int i = 0; i < workbookOld.getNumberOfSheets(); i++) {
            sheetOld = workbookOld.getSheetAt(i);
            sheetNew = workbookNew.createSheet(sheetOld.getSheetName());
            this.transformSheet(workbookOld, workbookNew, sheetOld, sheetNew);
        }
    }

    private void transformSheet(Workbook workbookOld, Workbook workbookNew,
                                Sheet sheetOld, Sheet sheetNew) {

        sheetNew.setDisplayFormulas(sheetOld.isDisplayFormulas());
        sheetNew.setDisplayGridlines(sheetOld.isDisplayGridlines());
        sheetNew.setDisplayGuts(sheetOld.getDisplayGuts());
        sheetNew.setDisplayRowColHeadings(sheetOld.isDisplayRowColHeadings());
        sheetNew.setDisplayZeros(sheetOld.isDisplayZeros());
        sheetNew.setFitToPage(sheetOld.getFitToPage());
        sheetNew.setHorizontallyCenter(sheetOld.getHorizontallyCenter());
        sheetNew.setMargin(Sheet.BottomMargin,
                sheetOld.getMargin(Sheet.BottomMargin));
        sheetNew.setMargin(Sheet.FooterMargin,
                sheetOld.getMargin(Sheet.FooterMargin));
        sheetNew.setMargin(Sheet.HeaderMargin,
                sheetOld.getMargin(Sheet.HeaderMargin));
        sheetNew.setMargin(Sheet.LeftMargin,
                sheetOld.getMargin(Sheet.LeftMargin));
        sheetNew.setMargin(Sheet.RightMargin,
                sheetOld.getMargin(Sheet.RightMargin));
        sheetNew.setMargin(Sheet.TopMargin, sheetOld.getMargin(Sheet.TopMargin));
        sheetNew.setPrintGridlines(sheetNew.isPrintGridlines());
        sheetNew.setRightToLeft(sheetNew.isRightToLeft());
        sheetNew.setRowSumsBelow(sheetNew.getRowSumsBelow());
        sheetNew.setRowSumsRight(sheetNew.getRowSumsRight());
        sheetNew.setVerticallyCenter(sheetOld.getVerticallyCenter());

        Row rowNew;
        for (Row row : sheetOld) {
            rowNew = sheetNew.createRow(row.getRowNum());
            if (rowNew != null)
                this.transformRow(workbookOld, workbookNew, (XSSFRow) row, rowNew);
        }

        for (int i = 0; i < this.lastColumn; i++) {
            sheetNew.setColumnWidth(i, sheetOld.getColumnWidth(i));
            sheetNew.setColumnHidden(i, sheetOld.isColumnHidden(i));
        }

        for (int i = 0; i < sheetOld.getNumMergedRegions(); i++) {
            CellRangeAddress merged = sheetOld.getMergedRegion(i);
            sheetNew.addMergedRegion(merged);
        }
    }

    private void transformRow(Workbook workbookOld, Workbook workbookNew,
                              XSSFRow rowOld, Row rowNew) {
        Cell cellNew;
        rowNew.setHeight(rowOld.getHeight());
        for (Cell cell : rowOld) {
            cellNew = rowNew.createCell(cell.getColumnIndex(),
                    cell.getCellType());
            if (cellNew != null)
                this.transformCell(workbookOld, workbookNew, (XSSFCell) cell,
                        cellNew);
        }
        this.lastColumn = Math.max(this.lastColumn, rowOld.getLastCellNum());
    }

    private void transformCell(Workbook workbookOld, Workbook workbookNew,
                               XSSFCell cellOld, Cell cellNew) {
        cellNew.setCellComment(cellOld.getCellComment());

        Integer hash = cellOld.getCellStyle().hashCode();
        if (this.styleMap != null && !this.styleMap.containsKey(hash)) {
            this.transformStyle(workbookOld, workbookNew, hash, cellOld.getCellStyle(), workbookNew.createCellStyle());
        }
        if (this.styleMap != null) {
            cellNew.setCellStyle(this.styleMap.get(hash));
        }
        switch (cellOld.getCellType()) {
            case BOOLEAN:
                cellNew.setCellValue(cellOld.getBooleanCellValue());
                break;
            case ERROR:
                cellNew.setCellValue(cellOld.getErrorCellValue());
                break;
            case FORMULA:
                cellNew.setCellValue(cellOld.getCellFormula());
                break;
            case NUMERIC:
                cellNew.setCellValue(cellOld.getNumericCellValue());
                break;
            case STRING:
                cellNew.setCellValue(cellOld.getStringCellValue());
                break;
            default:
                break;
        }
    }

    private void transformStyle(Workbook workbookOld, Workbook workbookNew,
                                Integer hash, XSSFCellStyle styleOld, CellStyle styleNew) {
        try {
            styleNew.setAlignment(styleOld.getAlignment());
            styleNew.setBorderBottom(styleOld.getBorderBottom());
            styleNew.setBorderLeft(styleOld.getBorderLeft());
            styleNew.setBorderRight(styleOld.getBorderRight());
            styleNew.setBorderTop(styleOld.getBorderTop());
            styleNew.setDataFormat(this.transformDataFormat(workbookOld, workbookNew, styleOld.getDataFormat()));
            styleNew.setFillBackgroundColor(styleOld.getFillBackgroundColor());
            styleNew.setFillForegroundColor(styleOld.getFillForegroundColor());
            styleNew.setFillPattern(styleOld.getFillPattern());
            styleNew.setFont(this.transformFont(workbookNew, styleOld.getFont()));
            styleNew.setHidden(styleOld.getHidden());
            styleNew.setIndention(styleOld.getIndention());
            styleNew.setLocked(styleOld.getLocked());
            styleNew.setVerticalAlignment(styleOld.getVerticalAlignment());
            styleNew.setWrapText(styleOld.getWrapText());
        } catch (Exception e) {
            styleNew.setAlignment(HorizontalAlignment.CENTER);
            styleNew.setBorderBottom(BorderStyle.THIN);
            styleNew.setBorderLeft(BorderStyle.THIN);
            styleNew.setBorderRight(BorderStyle.THIN);
            styleNew.setBorderTop(BorderStyle.THIN);
            styleNew.setDataFormat(workbookNew.createDataFormat().getFormat("General"));
            styleNew.setVerticalAlignment(VerticalAlignment.CENTER);
        }
        this.styleMap.put(hash, styleNew);
    }

    private short transformDataFormat(Workbook workbookOld, Workbook workbookNew,
                                      short index) {
        DataFormat formatOld = workbookOld.createDataFormat();
        DataFormat formatNew = workbookNew.createDataFormat();
        String format = formatOld.getFormat(index);
        if (format == null) {
            format = "General";
        }
        return formatNew.getFormat(format);
    }

    private Font transformFont(Workbook workbookNew, XSSFFont fontOld) {
        Font fontNew = workbookNew.createFont();
        fontNew.setBold(fontOld.getBold());
        fontNew.setCharSet(fontOld.getCharSet());
        fontNew.setColor(fontOld.getColor());
        fontNew.setFontName(fontOld.getFontName());
        fontNew.setFontHeight(fontOld.getFontHeight());
        fontNew.setItalic(fontOld.getItalic());
        fontNew.setStrikeout(fontOld.getStrikeout());
        fontNew.setTypeOffset(fontOld.getTypeOffset());
        fontNew.setUnderline(fontOld.getUnderline());
        return fontNew;
    }
}