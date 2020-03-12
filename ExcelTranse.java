import org.apache.commons.io.FileUtils;
import org.apache.commons.io.filefilter.TrueFileFilter;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import smartsuite.app.common.excel.ExcelCopyUtil;

import javax.servlet.ServletOutputStream;
import java.io.*;
import java.util.*;

public class ExcelTranse {



    public void excelTranse(File file ,String transeFilePath) throws Exception{


        InputStream inp;

        try {
            inp = new FileInputStream(file.getPath());
            Workbook wb = WorkbookFactory.create(inp);

            Workbook newWb = new XSSFWorkbook();
            Sheet copia = newWb.createSheet();
            Sheet sheet = wb.getSheetAt(0);
            Iterator<Row> rows = sheet.iterator();

            while(rows.hasNext()){
                Row row = rows.next();
                Row newRow = copia.createRow(row.getRowNum());
                Iterator<Cell> cells = row.cellIterator();
                while( cells.hasNext()){
                    Cell cell = cells.next();
                    Cell newCell =  newRow.createCell(cell.getColumnIndex());
                    int type = cell.getCellType();
                    List<CellStyle> styleList = new ArrayList<CellStyle>();

                    CellStyle newCellStyle = (XSSFCellStyle) getSameCellStyle(cell, newCell, styleList);

                    newCell.setCellStyle(newCellStyle);

                    switch (cell.getCellType()) {
                        case 1:
                            newCell.setCellValue(cell.getStringCellValue());
                            break;
                        case 0:
                            newCell.setCellValue(cell.getNumericCellValue());
                            break;
                        case 3:
                            newCell.setCellType(CellType.BLANK);
                            break;
                        case 4:
                            newCell.setCellValue(cell.getBooleanCellValue());
                            break;
                        case 5:
                            newCell.setCellErrorValue(cell.getErrorCellValue());
                            break;
                        case 2:
                            newCell.setCellFormula(cell.getCellFormula());
                            formulaInfoList.add(new FormulaInfo(cell.getSheet().getSheetName(), cell.getRowIndex(), cell
                                    .getColumnIndex(), cell.getCellFormula()));
                            break;
                        default:
                            break;
                    }


                }

            }


            OutputStream out = new FileOutputStream(new File(transeFilePath));
            newWb.write(out);
            out.close();

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
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


    @Test
    public void readMapperDir(){


        String repLine = "";
        int buffer;
        String isSurveyDir = "C:\\project\\survey";

        String isAsisDir = isSurveyDir+File.separator+"as-is";
        String isTobeDir = isSurveyDir+File.separator+"to-be";


        Map<String,Map<String,Object>> oracleMapperMap = new HashMap<String, Map<String, Object>>();
        Map<String,Map<String,Object>> hanaMapperMap = new HashMap<String, Map<String, Object>>();

        try{

            for (File info : FileUtils.listFilesAndDirs(new File(isAsisDir), TrueFileFilter.INSTANCE, TrueFileFilter.INSTANCE)) {
                if(info.isDirectory()) {

                    ArrayList<File> getFiles = new ArrayList<File>();

                    this.arrayDirInFileList(getFiles, info);

                    for (File inDirIsFile : getFiles) {

                        String toBeFileName = inDirIsFile.getName();
                        if(toBeFileName.indexOf("xlsx") == -1 && toBeFileName.indexOf("xls") > -1){
                            toBeFileName = toBeFileName.replaceAll("xls","xlsx");
                        }

                       this.excelTranse(inDirIsFile , isTobeDir+File.separator+toBeFileName);
                    }
                }
            }



            System.out.println("check");

        }catch (Exception e){
            System.out.println(e.getMessage());
        }
    }


    public ArrayList<File> arrayDirInFileList(ArrayList<File> files , File dir){

        if(dir.isDirectory()) {
            File[] children = dir.listFiles();
            for(File f : children) {
                // 재귀 호출 사용
                // 하위 폴더 탐색 부분
                arrayDirInFileList(files,f);
            }
        } else {
            files.add(dir);
        }

        return files;
    }


    public Map<String,Object> readMapperFile(String filePath){


        String repLine = "";
        int buffer;
        Map<String,Object> resultMapper = new HashMap<String, Object>();
        String line = "";
        try{
            // read
            File file = new File(filePath);
            BufferedReader bufferedReader = new BufferedReader(new FileReader(file));
            while ((line = bufferedReader.readLine()) != null) {

                if(line.indexOf("id=\"") > -1){ //id=를 찾는다.

                    int startIdText = line.indexOf("id=\"");
                    int startIdTextEnd = line.indexOf("\"",startIdText)+1;
                    int endIdText = line.indexOf("\"",startIdTextEnd);
                    //resultMapper.put()
                    String idText = line.substring(startIdTextEnd,endIdText);

                    resultMapper.put(idText,idText);
                }

            }

            bufferedReader.close();
        }catch (Exception e){
            System.out.println(e.getMessage());
        }


        return resultMapper;
    }


}
