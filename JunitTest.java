
import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.ss.formula.functions.Columns;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import org.springframework.core.io.ClassPathResource;

import javax.swing.*;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

public class JunitTest {

    //@Test
    /*public void test() throws Exception {
        List<String> headers = new ArrayList<String>(Arrays.asList("Param 1", "Param 2", "Param 3", "Param 4", "Param 5"));
        List<List<Object>> rows = new ArrayList<>();
        for(int i = 0; i < 10; i++){
            List<Object> row = new ArrayList<>();
            for(int k = 0; k < 5; k++){
                row.add(String.format("Val(%s,%s)", i, k));
            }
            rows.add(row);
        }

        // loading areas and commands using XmlAreaBuilder
        try(InputStream is = new FileInputStream(new File("C:\\solution\\testfile\\dynamic_columns_demo.xls"))) {
            try (OutputStream os = new FileOutputStream(new File("C:\\solution\\testfile\\dynamic_columns_demo_out.xls"))) {
                Transformer transformer = TransformerFactory.createTransformer(is, os);
                AreaBuilder areaBuilder = new XmlAreaBuilder(configInputStream, transformer);
                List<Area> xlsAreaList = areaBuilder.build();
                Area xlsArea = xlsAreaList.get(0);
                // creating context
                Context context = PoiTransformer.createInitialContext();
                context.putVar("headers", headers);
                context.putVar("rows", rows);
                // applying transformation
                logger.info("Applying area " + xlsArea.getAreaRef() + " at cell " + new CellRef("Result!A1"));
                xlsArea.applyAt(new CellRef("Result!A1"), context);
                // saving the results to file
                transformer.write();
                logger.info("Complete");
            }
        }
    }



    }*/




    @Test
    public void sxssfDynamicColumns() throws Exception {
        List<Map<String, Object>> lotsOfStuff = createLotsOfStuff();
/*
        Map<String, SheetData> sheetMap = new LinkedHashMap();
        Context context = new PoiContext();
        context.putVar("lotsOStuff", lotsOfStuff);
        context.putVar("columns", new Columns());

        //Workbook template = WorkbookFactory.create(resource.getInputStream());

        try(InputStream in = new FileInputStream(new File("C:\\solution\\testfile\\dynamic_columns_demo.xls"))) {
            try (OutputStream os = new FileOutputStream(new File("C:\\solution\\testfile\\dynamic_columns_demo_out.xls"))) {
               *//* Workbook workbook = WorkbookFactory.create(in);
                PoiTransformer transformer = PoiTransformer.createSxssfTransformer(workbook, 5, false);

                AreaBuilder areaBuilder;
                areaBuilder = new XlsCommentAreaBuilder(transformer);
                List<Area> xlsAreaList = areaBuilder.build();
                Area xlsArea = xlsAreaList.get(0);
                xlsArea.applyAt(new CellRef("Result!A1"), context);
                SXSSFWorkbook workbook2 = (SXSSFWorkbook) transformer.getWorkbook();

                //todo sxssf는 cell에 data를 넣으면 flush를 해줘야하는데 해당 내역이 필요없는건지.. jxls 와 합치면..?
                workbook2.write(os);

                //SXSSFWorkbook은 sheet를 메모리에 저장하는게 있어서 날려야하는것으로 보인다.
                workbook2.dispose();*//*

                Workbook workbook = WorkbookFactory.create(in);
               // PoiTransformer transformer = PoiTransformer.createSxssfTransformer(workbook, 5, false);

                int numberOfSheets =  workbook.getNumberOfSheets();
                for(int i = 0; i < numberOfSheets; ++i) {
                     Sheet sheet = workbook.getSheetAt(i);
                     SheetData sheetData = PoiSheetData.createSheetData(sheet, null);
                     sheetMap.put(sheetData.getSheetName(), sheetData);
                 }

            }*/
        //}


        /**
         * private void readCellData() {
         *         int numberOfSheets = this.workbook.getNumberOfSheets();
         *
         *         for(int i = 0; i < numberOfSheets; ++i) {
         *             Sheet sheet = this.workbook.getSheetAt(i);
         *             SheetData sheetData = PoiSheetData.createSheetData(sheet, this);
         *             this.sheetMap.put(sheetData.getSheetName(), sheetData);
         *         }
         *
         *     }
         *
         */

        /**
         *  public static PoiSheetData createSheetData(Sheet sheet, PoiTransformer transformer) {
         *         PoiSheetData sheetData = new PoiSheetData();
         *         sheetData.setTransformer(transformer);
         *         sheetData.sheet = sheet;
         *         sheetData.sheetName = sheet.getSheetName();
         *         int numberOfRows = sheet.getLastRowNum() + 1;
         *         int numberOfColumns = -1;
         *
         *         int i;
         *         for(i = 0; i < numberOfRows; ++i) {
         *             RowData rowData = PoiRowData.createRowData(sheet.getRow(i), transformer);
         *             sheetData.rowDataList.add(rowData);
         *             if (rowData != null && rowData.getNumberOfCells() > numberOfColumns) {
         *                 numberOfColumns = rowData.getNumberOfCells();
         *             }
         *         }
         *
         *         for(i = 0; i < sheet.getNumMergedRegions(); ++i) {
         *             CellRangeAddress region = sheet.getMergedRegion(i);
         *             sheetData.mergedRegions.add(region);
         *         }
         *
         *         if (numberOfColumns > 0) {
         *             sheetData.columnWidth = new int[numberOfColumns];
         *
         *             for(i = 0; i < numberOfColumns; ++i) {
         *                 sheetData.columnWidth[i] = sheet.getColumnWidth(i);
         *             }
         *         }
         *
         *         return sheetData;
         *     }
         *
         *
         */
        Workbook workbook = new SXSSFWorkbook(200);

        /**
         *
         * dropdown
         *
         * Workbook workbook = new SXSSFWorkbook(200);
         *
         *
         *
         * public static void createDropDown(String[] status, String defaultValue, Row row, int rowid, int cellid) {
         *
         * Cell cell = row.createCell(cellid);
         * Sheet sheet = row.getSheet();
         * DataValidationHelper dvHelper = sheet.getDataValidationHelper();
         * XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint) dvHelper.createExplicitListConstraint(status);
         * CellRangeAddressList addressList = new CellRangeAddressList(rowid, rowid, cellid, cellid);
         * XSSFDataValidation validation= (XSSFDataValidation) dvHelper.createValidation(dvConstraint, addressList);
         * validation.createErrorBox(ERROR_MESSSAGE, INVALID_DATA);
         * validation.setShowErrorBox(true);
         * sheet.addValidationData(validation);
         * sheet.getRow(rowid).getCell(cellid).setCellValue(defaultValue);
         * }
         *
         *
         *
         */


        /**
         * /**
         *  * Create a drop-down menu in excel using jxls
         *  * @param sheet
         *   * @param y The line number of the drop-down menu to be created (starting with 0)
         *   * @param x The column number of the drop-down menu to be created (starting with 0)
         *  * @return
         *
         * public void createListBox (Sheet sheet,int y, int x){
         * 		//Generate a drop - down list
         * CellRangeAddressList
            regions = new CellRangeAddressList(y, y, x, x);//CellRangeAddressList(int firstRow,int lastRow,int firstCol,int lastCol);
         * 		/ /Generate the drop - down box content
         *DVConstraint constraint = DVConstraint.createExplicitListConstraint(new String[]{"Yes", "No"});
         *        //Add a drop-down menu to the sheet
         *HSSFDataValidation data_validation = new HSSFDataValidation(regions, constraint);
         *sheet.addValidationData(data_validation);
         *}
         *
         */
    }

    private List<Map<String, Object>> createLotsOfStuff() {
        Map<String, Object> stuff1 = new LinkedHashMap<String, Object>();
        Map<String, Object> stuff2 = new LinkedHashMap<String, Object>();

        stuff1.put("header0", "stuff_1_value0");
        stuff1.put("header1_dynamic", "stuff_1_value1");
        stuff1.put("header2_dynamic", "stuff_1_value2");
        stuff1.put("header3_dynamic", "stuff_1_value3");

        stuff2.put("header0", "stuff_2_value0");
        stuff2.put("header1_dynamic", "stuff_2_value1");
        stuff2.put("header2_dynamic", "stuff_2_value2");
        stuff2.put("header3_dynamic", "stuff_2_value3");

        return Arrays.asList(stuff1, stuff2);
    }


}
