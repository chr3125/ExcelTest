import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;
import org.junit
        .Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.*;


public class ExcelReader {

    private static void copySheet(XSSFSheet source, XSSFSheet destination){
        copySheet(source, destination, true);
    }

    private static void copySheet(XSSFSheet source, XSSFSheet destination, boolean copyStyle) {
        int maxColumnNum = 0;
        List<CellStyle> styleMap2 = (copyStyle) ? new ArrayList<CellStyle>() : null;
        for (int i = source.getFirstRowNum(); i <= source.getLastRowNum(); i++) {
            XSSFRow srcRow = source.getRow(i);
            XSSFRow destRow = destination.createRow(i);
            if (srcRow != null) {
                copyRow(source, destination, srcRow, destRow, styleMap2);
                if (srcRow.getLastCellNum() > maxColumnNum) {
                    maxColumnNum = srcRow.getLastCellNum();
                }
            }
        }
        for (int i = 0; i <= maxColumnNum; i++) {
            destination.autoSizeColumn(i);
            destination.setColumnWidth(i, source.getColumnWidth(i));
        }
    }


    private static void copyRow(XSSFSheet srcSheet, XSSFSheet destSheet, XSSFRow srcRow, XSSFRow destRow,
                                List<CellStyle> styleMap) {
        Set<CellRangeAddressWrapper> mergedRegions = new TreeSet<CellRangeAddressWrapper>();
        short dh = srcSheet.getDefaultRowHeight();
        if (srcRow.getHeight() != dh) {
            destRow.setHeight(srcRow.getHeight());
        }


        int j = srcRow.getFirstCellNum();
        if (j < 0) {
            j = 0;
        }
        for (; j <= srcRow.getLastCellNum(); j++) {
            XSSFCell oldCell = srcRow.getCell(j); // ancienne cell
            XSSFCell newCell = destRow.getCell(j); // new cell
            if (oldCell != null) {
                if (newCell == null) {
                    newCell = destRow.createCell(j);
                }

                copyCell(oldCell, newCell, styleMap);
                CellRangeAddress mergedRegion = getMergedRegion(srcSheet, srcRow.getRowNum(),
                        (short) oldCell.getColumnIndex());

                if (mergedRegion != null) {
                    CellRangeAddress newMergedRegion = new CellRangeAddress(mergedRegion.getFirstRow(),
                            mergedRegion.getLastRow(), mergedRegion.getFirstColumn(), mergedRegion.getLastColumn());
                    CellRangeAddressWrapper wrapper = new CellRangeAddressWrapper(newMergedRegion);
                    if (isNewMergedRegion(wrapper, mergedRegions)) {
                        mergedRegions.add(wrapper);
                        destSheet.addMergedRegion(wrapper.range);
                    }
                }
            }
        }
    }

    public static class CellRangeAddressWrapper implements Comparable<CellRangeAddressWrapper> {

        public CellRangeAddress range;

        public CellRangeAddressWrapper(CellRangeAddress theRange) {
            this.range = theRange;
        }

        public int compareTo(CellRangeAddressWrapper o) {

            if (range.getFirstColumn() < o.range.getFirstColumn()
                    && range.getFirstRow() < o.range.getFirstRow()) {
                return -1;
            } else if (range.getFirstColumn() == o.range.getFirstColumn()
                    && range.getFirstRow() == o.range.getFirstRow()) {
                return 0;
            } else {
                return 1;
            }

        }

    }

    private static void copyCell(Cell oldCell, Cell newCell, List<CellStyle> styleList) {
        if (styleList != null) {
            if (oldCell.getSheet().getWorkbook() == newCell.getSheet().getWorkbook()) {
                newCell.setCellStyle(oldCell.getCellStyle());
            } else {
                XSSFCellStyle newCellStyle = (XSSFCellStyle) getSameCellStyle(oldCell, newCell, styleList);
                if (newCellStyle == null) {
                    newCellStyle = (XSSFCellStyle) newCell.getSheet().getWorkbook().createCellStyle();
                    newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
                    styleList.add(newCellStyle);
                }
                newCell.setCellStyle(newCellStyle);

                /*

                DataFormat newDataFormat = newCell.getSheet().getWorkbook().createDataFormat();

                XSSFCellStyle newCellStyle = (XSSFCellStyle) getSameCellStyle(oldCell, newCell, styleList);
                if (newCellStyle == null) {
                    // Create a new cell style
                    XSSFFont oldFont = (XSSFFont) oldCell.getSheet().getWorkbook().getFontAt(oldCell.getCellStyle().getFontIndex());

                     *//*XSSFFont newFont = (XSSFFont) newCell
                            .getSheet()
                            .getWorkbook()
                            .findFont(oldFont.getBold(), oldFont.getColor(), oldFont.getFontHeight(),
                                    oldFont.getFontName(), oldFont.getItalic(), oldFont.getStrikeout(),
                                    oldFont.getTypeOffset(), oldFont.getUnderline());
                    if (newFont == null) {*//*
                    XSSFFont newFont = (XSSFFont) newCell.getSheet().getWorkbook().createFont();
                        newFont.setBold(oldFont.getBold());
                        //newFont.setColor(oldFont.getColor());
                        newFont.setFontHeight(oldFont.getFontHeight());
                        newFont.setFontName(oldFont.getFontName());
                        newFont.setItalic(oldFont.getItalic());
                        newFont.setStrikeout(oldFont.getStrikeout());
                        newFont.setTypeOffset(oldFont.getTypeOffset());
                        newFont.setUnderline(oldFont.getUnderline());
                        newFont.setCharSet(oldFont.getCharSet());
                        newFont.setThemeColor(oldFont.getThemeColor());
                        newFont.setColor(oldFont.getXSSFColor());


                    short newFormat = newDataFormat.getFormat(oldCell.getCellStyle().getDataFormatString());
                    newCellStyle = (XSSFCellStyle) newCell.getSheet().getWorkbook().createCellStyle();
                    newCellStyle.setFont(newFont);
                    newCellStyle.setDataFormat(newFormat);

                    newCellStyle.setAlignment(oldCell.getCellStyle().getAlignment());
                    newCellStyle.setHidden(oldCell.getCellStyle().getHidden());
                    newCellStyle.setLocked(oldCell.getCellStyle().getLocked());
                    newCellStyle.setWrapText(oldCell.getCellStyle().getWrapText());
                    newCellStyle.setBorderBottom(oldCell.getCellStyle().getBorderBottom());
                    newCellStyle.setBorderLeft(oldCell.getCellStyle().getBorderLeft());
                    newCellStyle.setBorderRight(oldCell.getCellStyle().getBorderRight());
                    newCellStyle.setBorderTop(oldCell.getCellStyle().getBorderTop());
                    newCellStyle.setBottomBorderColor(oldCell.getCellStyle().getBottomBorderColor());
                    newCellStyle.setFillBackgroundColor(oldCell.getCellStyle().getFillBackgroundColor());
                    newCellStyle.setFillForegroundColor(oldCell.getCellStyle().getFillForegroundColor());
                    newCellStyle.setFillPattern(oldCell.getCellStyle().getFillPattern());
                    newCellStyle.setIndention(oldCell.getCellStyle().getIndention());
                    newCellStyle.setLeftBorderColor(oldCell.getCellStyle().getLeftBorderColor());
                    newCellStyle.setRightBorderColor(oldCell.getCellStyle().getRightBorderColor());
                    newCellStyle.setRotation(oldCell.getCellStyle().getRotation());
                    newCellStyle.setTopBorderColor(oldCell.getCellStyle().getTopBorderColor());
                    newCellStyle.setVerticalAlignment(oldCell.getCellStyle().getVerticalAlignment());

                    styleList.add(newCellStyle);
                }
                newCell.setCellStyle(newCellStyle);*/
            }
        }
        switch (oldCell.getCellType()) {
            case 2:
                newCell.setCellValue(oldCell.getStringCellValue());
                break;
            case 1:
                newCell.setCellValue(oldCell.getNumericCellValue());
                break;
            case 4:
                newCell.setCellType(CellType.BLANK);
                break;
            case 5:
                newCell.setCellValue(oldCell.getBooleanCellValue());
                break;
            case 6:
                newCell.setCellErrorValue(oldCell.getErrorCellValue());
                break;
            case 3:
                newCell.setCellFormula(oldCell.getCellFormula());
                formulaInfoList.add(new FormulaInfo(oldCell.getSheet().getSheetName(), oldCell.getRowIndex(), oldCell
                        .getColumnIndex(), oldCell.getCellFormula()));
                break;
            default:
                break;
        }
    }
    private static class FormulaInfo {

        private String sheetName;
        private Integer rowIndex;
        private Integer cellIndex;
        private String formula;

        private FormulaInfo(String sheetName, Integer rowIndex, Integer cellIndex, String formula) {
            this.sheetName = sheetName;
            this.rowIndex = rowIndex;
            this.cellIndex = cellIndex;
            this.formula = formula;
        }

        public String getSheetName() {
            return sheetName;
        }

        public void setSheetName(String sheetName) {
            this.sheetName = sheetName;
        }

        public Integer getRowIndex() {
            return rowIndex;
        }

        public void setRowIndex(Integer rowIndex) {
            this.rowIndex = rowIndex;
        }

        public Integer getCellIndex() {
            return cellIndex;
        }

        public void setCellIndex(Integer cellIndex) {
            this.cellIndex = cellIndex;
        }

        public String getFormula() {
            return formula;
        }

        public void setFormula(String formula) {
            this.formula = formula;
        }
    }

    static List<FormulaInfo> formulaInfoList = new ArrayList<FormulaInfo>();

    public static void refreshFormula(XSSFWorkbook workbook) {
        for (FormulaInfo formulaInfo : formulaInfoList) {
            workbook.getSheet(formulaInfo.getSheetName()).getRow(formulaInfo.getRowIndex())
                    .getCell(formulaInfo.getCellIndex()).setCellFormula(formulaInfo.getFormula());
        }
        formulaInfoList.removeAll(formulaInfoList);
    }

    private static CellStyle getSameCellStyle(Cell oldCell, Cell newCell, List<CellStyle> styleList) {
        CellStyle styleToFind = oldCell.getCellStyle();
        CellStyle currentCellStyle = null;
        CellStyle returnCellStyle = null;
        Iterator<CellStyle> iterator = styleList.iterator();

        XSSFFont oldFont = null;
        XSSFFont newFont = null;
        while (iterator.hasNext() && returnCellStyle == null) {
            currentCellStyle = iterator.next();

            if (currentCellStyle.getAlignment() != styleToFind.getAlignment()) {
                continue;
            }
            if (currentCellStyle.getHidden() != styleToFind.getHidden()) {
                continue;
            }
            if (currentCellStyle.getLocked() != styleToFind.getLocked()) {
                continue;
            }
            if (currentCellStyle.getWrapText() != styleToFind.getWrapText()) {
                continue;
            }
            if (currentCellStyle.getBorderBottom() != styleToFind.getBorderBottom()) {
                continue;
            }
            if (currentCellStyle.getBorderLeft() != styleToFind.getBorderLeft()) {
                continue;
            }
            if (currentCellStyle.getBorderRight() != styleToFind.getBorderRight()) {
                continue;
            }
            if (currentCellStyle.getBorderTop() != styleToFind.getBorderTop()) {
                continue;
            }
            if (currentCellStyle.getBottomBorderColor() != styleToFind.getBottomBorderColor()) {
                continue;
            }
            if (currentCellStyle.getFillBackgroundColor() != styleToFind.getFillBackgroundColor()) {
                continue;
            }
            if (currentCellStyle.getFillForegroundColor() != styleToFind.getFillForegroundColor()) {
                continue;
            }
            if (currentCellStyle.getFillPattern() != styleToFind.getFillPattern()) {
                continue;
            }
            if (currentCellStyle.getIndention() != styleToFind.getIndention()) {
                continue;
            }
            if (currentCellStyle.getLeftBorderColor() != styleToFind.getLeftBorderColor()) {
                continue;
            }
            if (currentCellStyle.getRightBorderColor() != styleToFind.getRightBorderColor()) {
                continue;
            }
            if (currentCellStyle.getRotation() != styleToFind.getRotation()) {
                continue;
            }
            if (currentCellStyle.getTopBorderColor() != styleToFind.getTopBorderColor()) {
                continue;
            }
            if (currentCellStyle.getVerticalAlignment() != styleToFind.getVerticalAlignment()) {
                continue;
            }
            oldFont = (XSSFFont) oldCell.getSheet().getWorkbook().getFontAt(oldCell.getCellStyle().getFontIndex());
            newFont = (XSSFFont) newCell.getSheet().getWorkbook().getFontAt(currentCellStyle.getFontIndex());

            if (newFont.getBold() == oldFont.getBold()) {
                continue;
            }
            if (newFont.getColor() == oldFont.getColor()) {
                continue;
            }
            if (newFont.getFontHeight() == oldFont.getFontHeight()) {
                continue;
            }
            if (newFont.getFontName() == oldFont.getFontName()) {
                continue;
            }
            if (newFont.getItalic() == oldFont.getItalic()) {
                continue;
            }
            if (newFont.getStrikeout() == oldFont.getStrikeout()) {
                continue;
            }
            if (newFont.getTypeOffset() == oldFont.getTypeOffset()) {
                continue;
            }
            if (newFont.getUnderline() == oldFont.getUnderline()) {
                continue;
            }
            if (newFont.getCharSet() == oldFont.getCharSet()) {
                continue;
            }
            if (oldCell.getCellStyle().getDataFormatString().equals(currentCellStyle.getDataFormatString())) {
                continue;
            }

            returnCellStyle = currentCellStyle;
        }
        return returnCellStyle;
    }

    public static CellRangeAddress getMergedRegion(XSSFSheet sheet, int rowNum, short cellNum) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress merged = sheet.getMergedRegion(i);
            if (merged.isInRange(rowNum, cellNum)) {
                return merged;
            }
        }
        return null;
    }
    private static boolean isNewMergedRegion(CellRangeAddressWrapper newMergedRegion,
                                             Set<CellRangeAddressWrapper> mergedRegions) {
        return !mergedRegions.contains(newMergedRegion);
    }

    private static void copyPrintTitle(Sheet newSheet, Sheet sheetToCopy) {
        int nbNames = sheetToCopy.getWorkbook().getNumberOfNames();
        Name name = null;
        String formula = null;

        String part1S = null;
        String part2S = null;
        String formS = null;
        String formF = null;
        String part1F = null;
        String part2F = null;
        int rowB = -1;
        int rowE = -1;
        int colB = -1;
        int colE = -1;

        for (int i = 0; i < nbNames; i++) {
            name = sheetToCopy.getWorkbook().getNameAt(i);
            if (name.getSheetIndex() == sheetToCopy.getWorkbook().getSheetIndex(sheetToCopy)) {
                if (name.getNameName().equals("Print_Titles")
                        || name.getNameName().equals(XSSFName.BUILTIN_PRINT_TITLE)) {
                    formula = name.getRefersToFormula();
                    int indexComma = formula.indexOf(",");
                    if (indexComma == -1) {
                        indexComma = formula.indexOf(";");
                    }
                    String firstPart = null;
                    ;
                    String secondPart = null;
                    if (indexComma == -1) {
                        firstPart = formula;
                    } else {
                        firstPart = formula.substring(0, indexComma);
                        secondPart = formula.substring(indexComma + 1);
                    }

                    formF = firstPart.substring(firstPart.indexOf("!") + 1);
                    part1F = formF.substring(0, formF.indexOf(":"));
                    part2F = formF.substring(formF.indexOf(":") + 1);

                    if (secondPart != null) {
                        formS = secondPart.substring(secondPart.indexOf("!") + 1);
                        part1S = formS.substring(0, formS.indexOf(":"));
                        part2S = formS.substring(formS.indexOf(":") + 1);
                    }

                    rowB = -1;
                    rowE = -1;
                    colB = -1;
                    colE = -1;
                    String rowBs, rowEs, colBs, colEs;
                    if (part1F.lastIndexOf("$") != part1F.indexOf("$")) {
                        rowBs = part1F.substring(part1F.lastIndexOf("$") + 1, part1F.length());
                        rowEs = part2F.substring(part2F.lastIndexOf("$") + 1, part2F.length());
                        rowB = Integer.parseInt(rowBs);
                        rowE = Integer.parseInt(rowEs);
                        if (secondPart != null) {
                            colBs = part1S.substring(part1S.lastIndexOf("$") + 1, part1S.length());
                            colEs = part2S.substring(part2S.lastIndexOf("$") + 1, part2S.length());
                            colB = CellReference.convertColStringToIndex(colBs);// CExportExcelHelperPoi.convertColumnLetterToInt(colBs);
                            colE = CellReference.convertColStringToIndex(colEs);// CExportExcelHelperPoi.convertColumnLetterToInt(colEs);
                        }
                    } else {
                        colBs = part1F.substring(part1F.lastIndexOf("$") + 1, part1F.length());
                        colEs = part2F.substring(part2F.lastIndexOf("$") + 1, part2F.length());
                        colB = CellReference.convertColStringToIndex(colBs);// CExportExcelHelperPoi.convertColumnLetterToInt(colBs);
                        colE = CellReference.convertColStringToIndex(colEs);// CExportExcelHelperPoi.convertColumnLetterToInt(colEs);

                        if (secondPart != null) {
                            rowBs = part1S.substring(part1S.lastIndexOf("$") + 1, part1S.length());
                            rowEs = part2S.substring(part2S.lastIndexOf("$") + 1, part2S.length());
                            rowB = Integer.parseInt(rowBs);
                            rowE = Integer.parseInt(rowEs);
                        }
                    }

                }
            }
        }
    }


    @Test
    public static void main(String[] args) throws IOException, InvalidFormatException {

        Workbook sourceWorkBook = new XSSFWorkbook(OPCPackage.open("C:\\solution\\testfile\\test_1.xlsx"));
        Workbook destinationWorkBook = new XSSFWorkbook(OPCPackage.open("C:\\solution\\testfile\\test_2.xlsx"));
        OutputStream out = new FileOutputStream(new File("C:\\solution\\testfile\\testFile_out.xlsx"));

      /*  wb_1.setSheetName(0, "Actual");
        wb_1.createSheet("Last Week");*/

        /*Get sheets from the temp file*/
        XSSFSheet destination = ((XSSFWorkbook) destinationWorkBook).getSheetAt(0);
        XSSFSheet source = ((XSSFWorkbook) sourceWorkBook).getSheetAt(0);

        copySheet(source, destination);

        destinationWorkBook.write(out);
        out.close();
      /*  OutputStream os = new FileOutputStream("C:\\solution\\testfile\\test_2.xlsx");
        sourceWorkBook.write(os);
        os.flush();
        os.close();
        sourceWorkBook.close();*/

        /*OutputStream os = new FileOutputStream("C:\\solution\\testfile\\hello.xlsx");

        System.out.print ("If you arrived here, it means you're good boy");
        wb_1.write(os);
        os.flush();
        os.close();
        wb_1.close();*/
    }
}
