package smartsuite.app.common.excel;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.stereotype.Service;
import smartsuite.app.bp.admin.mailWork.MailWorkController;
import smartsuite.app.common.cert.util.EdocStringUtil;
import smartsuite.app.common.excel.bean.CellInfoBean;
import smartsuite.app.common.excel.bean.ExcelInfoBean;
import smartsuite.app.common.excel.bean.RowInfoBean;
import smartsuite.app.common.excel.bean.SheetInfoBean;
import smartsuite.security.web.filter.advanced.XssAction;
import smartsuite.upload.core.entity.FileGroup;
import smartsuite.upload.core.entity.FileItem;
import smartsuite.upload.core.service.FileService;
import smartsuite.upload.core.store.ParameterizedLocation;
import smartsuite.upload.core.store.RootLocation;
import smartsuite.upload.spring.web.multipart.SimpleMultipartFileItem;

import javax.inject.Inject;
import java.io.*;
import java.util.*;

/**
 * Excel을 Copy하고, 해당 하는 Data / Style들에 대해서 기본값들을 정의하기 위하여 Create Util
 */
@SuppressWarnings("unused")
@Service
public class ExcelCreateUtil {

    static final Logger LOG = LoggerFactory.getLogger(ExcelCreateUtil.class);

    // value 로만 단일 값에 대해서 처리하기로함
    private static final String STARTEXPRESSIONTOKEN = "${";
    private static final String ENDEXPRESSIONTOKEN = "}";

    // list 처리할때 사용 하기로 함
    private static final String STARTFORMULATOKEN = "$[";
    private static final String ENDFORMULATOKEN = "]";

    private static final int REPLACE_TYPE_VALUE = 2;
    private static final int REPLACE_TYPE_LIST = 1;
    private static final int REPLACE_TYPE_NONE = 0;

    @Inject
    FileService fileService;

    @Inject
    ExcelReaderUtil excelReaderUtil;

    @Value ("#{file['file.upload.path']}")
    String fileUploadPath;


    /**
     * Excel을 Copy 시작하는 메소드 최초에 Sheet의 틀을 복사한다.
     * @param sheetList
     * @param destWorkbook
     */
    public static void copyExcel(List<SheetInfoBean> sheetList, XSSFWorkbook destWorkbook) {
        int maxColumnNum = 0;

        for (int i = 0; i < sheetList.size(); i++) {
            SheetInfoBean sheet = sheetList.get(i);
            XSSFSheet destSheet = destWorkbook.createSheet(sheet.getXls_work_sht_nm());

            copySheet(sheet.getRowList(), destSheet);
        }

    }

    /**
     * Sheet를 복사하여, 복사한 시트에 행 (ROW) 체크하여, Row 별로 Copy가 될수있도록 처리한다.
     * @param rowList
     * @param destination
     */
    public static void copySheet(List<RowInfoBean> rowList, XSSFSheet destination) {
        int maxColumnNum = 0;

        for (int i = 0; i < rowList.size(); i++) {
            RowInfoBean sheetRow = rowList.get(i);
            XSSFRow destRow = destination.createRow(sheetRow.getRow_no());
            copyRow( sheetRow, destRow);
        }
        for (int i = 0; i < rowList.size(); i++) {
            RowInfoBean sheetRow = rowList.get(i);
            destination.autoSizeColumn(i);
            // sheet별이 아닌 excel sheet 전체를 돌리게 끔하여서 가져와야할껏으로 보임.
            //destination.setColumnWidth(i, sheetRow.getColumn_width());
        }
    }


    /**
     * Row 내에 있는 Cell들을 취득하여, 복사 가능하도록 처리하는 메소드.
     * @param srcRow
     * @param destRow
     */
    private static void copyRow(  RowInfoBean srcRow, XSSFRow destRow) {
        //Set<CellRangeAddressWrapper> mergedRegions = new TreeSet<CellRangeAddressWrapper>();

        List<CellInfoBean> rowList = srcRow.getCellList();

        for (int i = 0; i < rowList.size(); i++) {
            CellInfoBean cellInfoBean = rowList.get(i);

            XSSFCell newCell = destRow.getCell(cellInfoBean.getCell_index()); // new cell
            if (newCell == null) {
                newCell = destRow.createCell(cellInfoBean.getCell_index());
            }

            copyCell(cellInfoBean, newCell);

            /** TODO
             * Merge된 Cell이 존재할 경우 아래의 주석을 해제하면 되는데, 별도의 Test가 선행되어야할 것으로 보임.
             * Excel Copy to Copy에서는 정상적이나, Data에 대한 치환시 문제가 생길것으로 보임.
             */
            /*CellRangeAddress mergedRegion = getMergedRegion(srcSheet, srcRow.getRowNum(),
                    (short) oldCell.getColumnIndex());

            if (mergedRegion != null) {
                CellRangeAddress newMergedRegion = new CellRangeAddress(mergedRegion.getFirstRow(),
                        mergedRegion.getLastRow(), mergedRegion.getFirstColumn(), mergedRegion.getLastColumn());
                CellRangeAddressWrapper wrapper = new CellRangeAddressWrapper(newMergedRegion);
                if (isNewMergedRegion(wrapper, mergedRegions)) {
                    mergedRegions.add(wrapper);
                    destSheet.addMergedRegion(wrapper.range);
                }
            }*/
        }
    }

    /**
     * Cell의 String Value내에 객체를 표기하는 형태가 존재하는지 찾고, 분류값을 지정한다. ( 해당 Method는 Row 내에 1개의 구분값(리스트,단일객체)가 있다는 전재하이다.)
     * -- List 객체의 경우  $[ ]
     * -- 단일 객체의 경우  ${ }
     * @param srcRow
     * @return
     */
    private static int checkCellListAndValue(  RowInfoBean srcRow) {
       boolean checkReplaceValue = false;
       boolean checkReplaceList = false;

       // 현재까지 기준으론 1개의 row에 단일 또는 리스트로만 정의하도록 한다.
       int checkReplaceType = 0;


        List<CellInfoBean> rowList = srcRow.getCellList();

        for (int i = 0; i < rowList.size(); i++) {
            CellInfoBean cellInfoBean = rowList.get(i);
            String getValue = cellInfoBean.getString_value() == null? "" : cellInfoBean.getString_value();

            if(!EdocStringUtil.isEmpty(getValue)){
                if (getValue.startsWith(ExcelCreateUtil.STARTEXPRESSIONTOKEN) && getValue.endsWith(ExcelCreateUtil.ENDEXPRESSIONTOKEN)) {
                    checkReplaceValue = true;
                    break;
                }else if(getValue.startsWith(ExcelCreateUtil.STARTFORMULATOKEN) && getValue.endsWith(ExcelCreateUtil.ENDFORMULATOKEN)){
                    checkReplaceList = true;
                    break;
                }
            }
        }


        if(checkReplaceList){  // 리스트
            checkReplaceType = ExcelCreateUtil.REPLACE_TYPE_LIST;
        }else if(checkReplaceValue){  // 단일
            checkReplaceType = ExcelCreateUtil.REPLACE_TYPE_VALUE;
        }else{
            checkReplaceType = ExcelCreateUtil.REPLACE_TYPE_NONE;
        }

        return checkReplaceType;
    }


    /**
     *  Merge Cell을 판별하기 위한 Method
     */
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

    /**
     * Cell의 Style 중 Color가 테마/기본/사용자 지정 등 컬러들이 여러가지가 있어, 이를 애초에 rgb 값으로 측정하기 위하여 별도로 Check 를 만듬.
     * @param rgb
     * @return
     */
    private  static XSSFColor getColorForRGB(int rgb) {
        java.awt.Color color =new java.awt.Color(rgb);

        XSSFColor xc = new XSSFColor(color);
        return xc;
    }

    /**
     * Cell Value / Info / Style 정보를 Copy 하기 위하여 각 객체를 가져오는 형태로 구현
     * @param oldCell
     * @param newCell
     */
    private static void copyCell(CellInfoBean oldCell, XSSFCell newCell) {

        XSSFFont newFont = (XSSFFont) newCell.getSheet().getWorkbook().createFont();

        newFont.setBold(oldCell.isBold());
        //newFont.setColor(oldFont.getColor());
        newFont.setFontHeight(oldCell.getFont_height());
        newFont.setFontName(oldCell.getFont_nm());
        newFont.setItalic(oldCell.isItalic());
        newFont.setStrikeout(oldCell.isStrikeout());
        newFont.setTypeOffset(oldCell.getType_offset());
        newFont.setUnderline(oldCell.getUnder_line());
        newFont.setCharSet(oldCell.getCharset());
        //newFont.setThemeColor(oldFont.getThemeColor());
        newFont.setColor(getColorForRGB(oldCell.getColor()));

        XSSFCellStyle newCellStyle = (XSSFCellStyle) newCell.getSheet().getWorkbook().createCellStyle();
        newCellStyle.setFont(newFont);
        newCellStyle.setDataFormat(oldCell.getDataformat());
        newCellStyle.setAlignment(HorizontalAlignment.forInt(oldCell.getAlignment_cd()));
        newCellStyle.setHidden(oldCell.isHidden());
        newCellStyle.setLocked(oldCell.isLocked());
        newCellStyle.setWrapText(oldCell.isWraptext());
        newCellStyle.setBorderBottom(BorderStyle.valueOf(oldCell.getBorder_bottom_cd()));
        newCellStyle.setBorderLeft(BorderStyle.valueOf(oldCell.getBorder_left_cd()));
        newCellStyle.setBorderRight(BorderStyle.valueOf(oldCell.getBorder_right_cd()));
        newCellStyle.setBorderTop(BorderStyle.valueOf(oldCell.getBorder_top_cd()));
        newCellStyle.setBottomBorderColor(getColorForRGB(oldCell.getBottom_border_color()));
        newCellStyle.setFillBackgroundColor(getColorForRGB(oldCell.getFill_background_color()));
        newCellStyle.setFillForegroundColor(getColorForRGB(oldCell.getFill_foreground_color()));
        newCellStyle.setFillPattern(FillPatternType.forInt(oldCell.getFill_pattern()));
        newCellStyle.setIndention(oldCell.getIndention());
        newCellStyle.setLeftBorderColor(getColorForRGB(oldCell.getLeft_border_color()));
        newCellStyle.setRightBorderColor(getColorForRGB(oldCell.getRight_border_color()));
        newCellStyle.setRotation(oldCell.getRotation());
        newCellStyle.setTopBorderColor(getColorForRGB(oldCell.getTop_border_color()));
        newCellStyle.setVerticalAlignment(VerticalAlignment.forInt(oldCell.getVertical_alignment_cd()));
        newCell.setCellValue(oldCell.getCell_type());
        newCell.setCellStyle(newCellStyle);
        newCell.getSheet().setColumnWidth(oldCell.getCell_index(),oldCell.getCol_width());



        switch (oldCell.getCell_type()) {
            case 1:
                newCell.setCellType(1);
                newCell.setCellValue(oldCell.getString_value());
                break;
            case 0:
                newCell.setCellType(0);
                newCell.setCellValue(oldCell.getDouble_value());
                break;
            case 3:
                newCell.setCellType(3);
                break;
            case 4:
                newCell.setCellType(4);
                newCell.setCellValue(oldCell.isBoolean_value());
                break;
            case 5:
                newCell.setCellType(5);
                newCell.setCellErrorValue(oldCell.getError_value());
                break;
            case 2:
                newCell.setCellType(2);
                newCell.setCellFormula(oldCell.getFormula_value());
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

    /**
     * Cell type을 측정하기 위하여 별도로 구성하였으나, 기본적인 형태의 Template 라면 별도로 구현은 하지 않아도 될것으로 판단.
     * @param workbook
     */
    public static void refreshFormula(XSSFWorkbook workbook) {
        for (FormulaInfo formulaInfo : formulaInfoList) {
            workbook.getSheet(formulaInfo.getSheetName()).getRow(formulaInfo.getRowIndex())
                    .getCell(formulaInfo.getCellIndex()).setCellFormula(formulaInfo.getFormula());
        }
        formulaInfoList.removeAll(formulaInfoList);
    }


    /**
     * Merge Cell 가져오기 Method
     * @param sheet
     * @param rowNum
     * @param cellNum
     * @return
     */
    public static CellRangeAddress getMergedRegion(XSSFSheet sheet, int rowNum, short cellNum) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress merged = sheet.getMergedRegion(i);
            if (merged.isInRange(rowNum, cellNum)) {
                return merged;
            }
        }
        return null;
    }
    /*private static boolean isNewMergedRegion(CellRangeAddressWrapper newMergedRegion,
                                             Set<CellRangeAddressWrapper> mergedRegions) {
        return !mergedRegions.contains(newMergedRegion);
    }*/


    /**
     * 이메일 업무 관리 화면에서 Template 만을 복사하려할때 사용하는 Method 첫번째 ( 1번째 row에 1번째/2번째 Cell에 선택하지 않으면 보이지 않는 Key값을 숨겨둔다. )
     * email_work_targ_id : mail set id / excel info 를 연결하는 Key
     * email_snd_log_id : 이메일 업무 발송 시 정보를 취득할수있는 Key ( mail send key )
     * 해당 메소드에서는 Sheet를 생성하고, 기본적인 1차원적인 값들을 저장한다.
     *
     * @param sheetList
     * @param destWorkbook
     * @param dataMapForSheet
     * @param headersList
     * @return
     */
    public static Map<String,Object> createExcelByListDataSetup(List<SheetInfoBean> sheetList, XSSFWorkbook destWorkbook,Map<String, Object> dataMapForSheet ,List<Map<String,Object>> headersList) {
        int maxColumnNum = 0;
        Map<String,Object> resultMap = new HashMap<String, Object>();

        for (int i = 0; i < sheetList.size(); i++) {
            SheetInfoBean sheet = sheetList.get(i);

            XSSFSheet destSheet = destWorkbook.createSheet(sheet.getXls_work_sht_nm());

            if(i==0){ //기본 정보 입력

                String emailWorkTargId = UUID.randomUUID().toString();
                XSSFRow destRow = destSheet.createRow(0);
                CellStyle rowStyle = destRow.getRowStyle();
                DataFormat newDataFormat = destSheet.getWorkbook().createDataFormat();

                XSSFCell emailWorkTargIdNewCell = destRow.createCell(1); // new cell
                CellInfoBean emailWorkTargIdCell  = new CellInfoBean();
                emailWorkTargIdCell.setString_value(emailWorkTargId); //email_work_targ_id
                emailWorkTargIdCell.setCell_type(1);
                emailWorkTargIdCell.setHidden(true);
                emailWorkTargIdCell.setBorder_bottom_cd((short)0);
                emailWorkTargIdCell.setBorder_left_cd((short)0);
                emailWorkTargIdCell.setBorder_right_cd((short)0);
                emailWorkTargIdCell.setBorder_top_cd((short)0);
                emailWorkTargIdCell.setDataformat(newDataFormat.getFormat(";;;"));
                copyCell(emailWorkTargIdCell, emailWorkTargIdNewCell);



                String emailSndLogId = UUID.randomUUID().toString();
                XSSFCell emailSndLogIdNewCell = destRow.createCell(2); // new cell
                CellInfoBean emailSndLogIdCell  = new CellInfoBean();
                emailSndLogIdCell.setString_value(emailSndLogId); //email_snd_log_id
                emailSndLogIdCell.setCell_type(1);
                copyCell(emailSndLogIdCell, emailSndLogIdNewCell);

                resultMap.put("email_work_targ_id",emailWorkTargId);
                resultMap.put("email_snd_log_id",emailSndLogId);
            }

            //sheet별로 데이터를 다르게 저장해놨음.
            Map<String, Object> dataListRow = (Map<String, Object>) dataMapForSheet.get(sheet.getXls_work_sht());
            createExcelForMapList(sheet.getRowList(), destSheet,dataListRow,headersList);
        }
        return resultMap;
    }


    /**
     *  넘어온 Data Map 에 맞게 객체들을 생성 + 치환하는 형태로 구현
     *   -> excel cell 내에 list 형이 있다면, data map list count 만큼 행을 생성하여, 치환
     *   -> excel cell 내에 단일 객체가 있다면, data map에서 해당 객체들을 find 하여, 치환
     *
     * @param rowList
     * @param destination
     * @param dataListRow
     * @param headersList
     */
    public static void createExcelForMapList(List<RowInfoBean> rowList, XSSFSheet destination , Map<String, Object> dataListRow , List<Map<String,Object>> headersList ) {
        int maxColumnNum = 0;
        int headerNextRow = 0;

        for (int i = 0; i < rowList.size(); i++) {
            RowInfoBean sheetRow = rowList.get(i);
            int rowNo = sheetRow.getRow_no();
            int rowCreateNum = 0;

            if(rowNo == 0) continue;

            int replaceType = checkCellListAndValue(sheetRow);

            if(replaceType == ExcelCreateUtil.REPLACE_TYPE_LIST){

                if (rowCreateNum == 0)  rowCreateNum = rowNo;

                int dumyCreateNum = sheetRow.getRow_no();
                String getRowNo = Integer.toString(sheetRow.getRow_no());
                Map<String,Map<String,Object>> dataList = (Map<String, Map<String, Object>>) dataListRow.get(getRowNo);

                int lastDataList = dataList.size();

                for(int b = 0; b < lastDataList; b++){
                    XSSFRow destRow = destination.createRow(dumyCreateNum);
                    copyRow(sheetRow, destRow);
                    dumyCreateNum++;
                }

                for(int a = 0; a < lastDataList; a++){
                    Map<String,Object> data = dataList.get(Integer.toString(rowCreateNum));
                    //data list 를 기준으로 copy 처리한다.
                    dataSetRowForMapList( destination.getRow(rowCreateNum) , data);
                    rowCreateNum++; // create할때마다 1을 증가시켜서 create하는 row가 계속 증식 가능하도록 처리한다.
                }

            }else if(replaceType == ExcelCreateUtil.REPLACE_TYPE_VALUE){

                if (rowCreateNum == 0)  rowCreateNum = rowNo;

                String getRowNo = Integer.toString(sheetRow.getRow_no());
                Map<String,Map<String,Object>> dataList = (Map<String, Map<String, Object>>) dataListRow.get(getRowNo);

                Map<String,Object> data = dataList.get(Integer.toString(rowCreateNum));

                XSSFRow destRow = destination.createRow(rowCreateNum);
                copyRow(sheetRow, destRow);

                //data list 를 기준으로 copy 처리한다.
                dataSetRowForMapValue( destination.getRow(rowCreateNum) , data);

            }else{
                if (rowCreateNum == 0)  rowCreateNum = rowNo;
                XSSFRow destRow = destination.createRow(rowCreateNum);
                copyRow(sheetRow, destRow);
            }


            /*if(headersList.size() > 0 && i == 0){
                XSSFRow destRow = destination.createRow(sheetRow.getRowNum());
                setHeadersRow(sheetRow, destRow,headersList);
                headerNextRow = sheetRow.getRowNum() + 1;*/

        }
        for (int i = 0; i < rowList.size(); i++) {
            RowInfoBean sheetRow = rowList.get(i);
            destination.autoSizeColumn(i);
            //destination.setColumnWidth(i, sheetRow.getColumn_width());
        }
    }


    /**
     * String value 내에 ${ }가 존재한다면, 이를 replace 처리하고, "." 이 객체안에 존재할 경우, 이를 key 값으로 보아 main key / sub key로 구분하여, 넘어온 data map에서 찾아 치환 하도록 한다.
     * 단일 객체 치환하는 Method
     * @param destRow
     * @param data
     */
    private static void dataSetRowForMapList(  XSSFRow destRow , Map<String,Object> data) {
        //Set<CellRangeAddressWrapper> mergedRegions = new TreeSet<CellRangeAddressWrapper>();

        for (int j = 0; j <= destRow.getLastCellNum(); j++) {
            XSSFCell getCell = destRow.getCell(j); // ancienne cell
            if(getCell == null) continue;

            String excelStringValue = getCell.getStringCellValue() == null? "" : getCell.getStringCellValue();
            excelStringValue = excelStringValue.replace(ExcelCreateUtil.STARTFORMULATOKEN,"");
            excelStringValue = excelStringValue.replace(ExcelCreateUtil.ENDFORMULATOKEN,"");

            String getValue = "";

            if(!StringUtils.isEmpty(excelStringValue)){
                String dataGetKey = "";

                if(excelStringValue.indexOf(".") > -1){ // $[ data.value ] 형식의 구조라면, 아래와 같이.

                    String[] splitValue = excelStringValue.split("\\.");
                    dataGetKey = splitValue[0];
                    excelStringValue = splitValue[1];
                    Map<String,Object> getValueMap = (Map<String,Object>)data.get(dataGetKey) == null? null : (Map<String,Object>)data.get(dataGetKey);
                    if(null != getValueMap){
                        getValue = getValueMap.get(excelStringValue) == null? "" : getValueMap.get(excelStringValue).toString();
                    }
                }else{ // $[ value ] 형식의 구조라면, 아래와 같이.
                    getValue = data.get(excelStringValue) == null? "" : data.get(excelStringValue).toString();
                }
            }

            if(StringUtils.isEmpty(getValue)){
                getValue = getCell.getStringCellValue() == null? "" : getCell.getStringCellValue();
            }

            getCell.setCellValue(getValue);
        }
    }


    /**
     * String value 내에  $[ ]가 존재한다면, 이를 replace 처리하고, "." 이 객체안에 존재할 경우, 이를 key 값으로 보아 main key / sub key로 구분하여, 넘어온 data map에서 찾아 치환 하도록 한다.
     * LIST 형에 대한 치환 처리 Method
     * @param destRow
     * @param data
     */
    private static void dataSetRowForMapValue(  XSSFRow destRow , Map<String,Object> data) {
        //Set<CellRangeAddressWrapper> mergedRegions = new TreeSet<CellRangeAddressWrapper>();

        for (int j = 0; j <= destRow.getLastCellNum(); j++) {
            XSSFCell getCell = destRow.getCell(j); // ancienne cell
            if(getCell == null) continue;

            String excelStringValue = getCell.getStringCellValue() == null? "" : getCell.getStringCellValue();
            excelStringValue = excelStringValue.replace(ExcelCreateUtil.STARTEXPRESSIONTOKEN,"");
            excelStringValue = excelStringValue.replace(ExcelCreateUtil.ENDEXPRESSIONTOKEN,"");

            String getValue = "";

            if(!StringUtils.isEmpty(excelStringValue)){
                String dataGetKey = "";

                if(excelStringValue.indexOf(".") > -1){ // ${ data.value } 형식의 구조라면, 아래와 같이.

                    String[] splitValue = excelStringValue.split("\\.");
                    dataGetKey = splitValue[0];
                    excelStringValue = splitValue[1];
                    Map<String,Object> getValueMap = (Map<String,Object>)data.get(dataGetKey) == null? null : (Map<String,Object>)data.get(dataGetKey);
                    if(null != getValueMap){
                        getValue = getValueMap.get(excelStringValue) == null? "" : getValueMap.get(excelStringValue).toString();
                    }
                }else{ // ${ value } 형식의 구조라면, 아래와 같이.
                    getValue = data.get(excelStringValue) == null? "" : data.get(excelStringValue).toString();
                }
            }

            if(StringUtils.isEmpty(getValue)){
                getValue = getCell.getStringCellValue() == null? "" : getCell.getStringCellValue();
            }

            getCell.setCellValue(getValue);
        }
    }

    /**
     * 추후 추가할 headers ( 헤더 영역을 동기적으로 부여가능하도록 하는 형태 )의 대한 임시 method
     * @param srcRow
     * @param destRow
     * @param headersList
     */
   /* private static void setHeadersRow(  RowInfoBean srcRow, XSSFRow destRow , List<Map<String,Object>> headersList) {
        //Set<CellRangeAddressWrapper> mergedRegions = new TreeSet<CellRangeAddressWrapper>();

        List<CellInfoBean> rowList = srcRow.getCellList();

        for (int i = 0; i < rowList.size(); i++) {
            CellInfoBean cellInfoBean = rowList.get(i);

            XSSFCell newCell = destRow.getCell(cellInfoBean.getCell_index()); // new cell
            if (newCell == null) {
                newCell = destRow.createCell(cellInfoBean.getCell_index());
            }

            *//*for(Map<String,Object> headers : headersList){

            }*//*


            copyCell(cellInfoBean, newCell);
        }
    }*/

    /**
     * 업무에서 이메일 견적을 생성할 시 추후 회신 시에 넘어오는 excel 파일이 정상적인지 판단하기 위함 및 cell index 위치 및 row의 치환을 위한 여러가지 방편으로 data map에 있는 list 및 각 객체를 위하여 발송전 template를 생성
     * 해당 파일은 tmp_form_file 로 저장되며, 회신 온 메일과 비교 분석하여, result map으로 나타나게 합니다.
     *
     * @param sheetList
     * @param dataListSheet
     * @param headersList
     * @return
     */
    public Map<String,Object> createExcelDataTemplateFirst(List<SheetInfoBean> sheetList, Map<String,Map<String,Map<String,Map<String,Object>>>> dataListSheet ,List<Map<String,Object>> headersList){

        Map<String,Object> resultMap = new HashMap<String, Object>();
        OutputStream out = null;
        InputStream io = null;
        XSSFWorkbook tempWorkBook = null;
        File file = null;
        try{

            String fileNm = UUID.randomUUID().toString() + ".xlsx";


            file = new File(fileUploadPath + UUID.randomUUID().toString() + FilenameUtils.EXTENSION_SEPARATOR + FilenameUtils.getExtension(fileNm));

            tempWorkBook = new XSSFWorkbook();

            out = new FileOutputStream(file);

            for (int i = 0; i < sheetList.size(); i++) {
                SheetInfoBean sheet = sheetList.get(i);


                XSSFSheet destSheet = tempWorkBook.createSheet(sheet.getXls_work_sht_nm());

                if(i==0){ //기본 정보 입력

                    String emailWorkTargId = UUID.randomUUID().toString();
                    XSSFRow destRow = destSheet.createRow(0);
                    CellStyle rowStyle = destRow.getRowStyle();
                    DataFormat newDataFormat = destSheet.getWorkbook().createDataFormat();

                    XSSFCell emailWorkTargIdNewCell = destRow.createCell(1); // new cell
                    CellInfoBean emailWorkTargIdCell  = new CellInfoBean();
                    emailWorkTargIdCell.setString_value(emailWorkTargId); //email_work_targ_id
                    emailWorkTargIdCell.setCell_type(1);
                    emailWorkTargIdCell.setHidden(true);
                    emailWorkTargIdCell.setBorder_bottom_cd((short)0);
                    emailWorkTargIdCell.setBorder_left_cd((short)0);
                    emailWorkTargIdCell.setBorder_right_cd((short)0);
                    emailWorkTargIdCell.setBorder_top_cd((short)0);
                    emailWorkTargIdCell.setDataformat(newDataFormat.getFormat(";;;"));
                    emailWorkTargIdCell.setCol_width(3);
                    copyCell(emailWorkTargIdCell, emailWorkTargIdNewCell);

                    String emailSndLogId = UUID.randomUUID().toString();
                    XSSFCell emailSndLogIdNewCell = destRow.createCell(2); // new cell
                    CellInfoBean emailSndLogIdCell  = new CellInfoBean();
                    emailSndLogIdCell.setString_value(emailSndLogId); //email_work_targ_id
                    emailSndLogIdCell.setCell_type(1);
                    emailSndLogIdCell.setHidden(true);
                    emailSndLogIdCell.setBorder_bottom_cd((short)0);
                    emailSndLogIdCell.setBorder_left_cd((short)0);
                    emailSndLogIdCell.setBorder_right_cd((short)0);
                    emailSndLogIdCell.setBorder_top_cd((short)0);
                    emailSndLogIdCell.setDataformat(newDataFormat.getFormat(";;;"));
                    emailSndLogIdCell.setCol_width(3);
                    copyCell(emailSndLogIdCell, emailSndLogIdNewCell);

                    resultMap.put("email_work_targ_id",emailWorkTargId);
                    resultMap.put("email_snd_log_id",emailSndLogId);
                }

                //sheet별로 데이터를 다르게 저장해놨음.
                Map<String,Map<String,Map<String,Object>>> dataListRow = dataListSheet.get(sheet.getXls_work_sht());
                createExcelForMapListTemplateCreate(sheet.getRowList(), destSheet,dataListRow,headersList);
            }


            tempWorkBook.write(out);

            io = FileUtils.openInputStream(file);

            String grpCd = UUID.randomUUID().toString();

            SimpleMultipartFileItem fileItem = new SimpleMultipartFileItem(UUID.randomUUID().toString(),grpCd, fileNm, FilenameUtils.getExtension(fileNm), file.length(), null, null, file);

            fileItem.setMultipartFile(new MockMultipartFile(fileNm, io));

            fileService.create(fileItem);

            resultMap.put("tmp_form_file",grpCd);  // template용 file
        }catch (Exception e){
            LOG.error(e.getMessage());
        }



        return resultMap;
    }


    /**
     * template 로 만들어둔 excel 파일의 변수명에 맞춰 넘어온 data map과 맞춰 치환 시키는 method
     * @param sheetList
     * @param destWorkbook
     * @param dataListSheet
     * @param headersList
     * @param templateCreationMap
     */
    public static  void createExcelDataSetup(List<SheetInfoBean> sheetList, XSSFWorkbook destWorkbook,Map<String,Map<String,Map<String,Map<String,Object>>>> dataListSheet ,List<Map<String,Object>> headersList , Map<String,Object> templateCreationMap){

        Map<String,Object> resultMap = new HashMap<String, Object>();
        OutputStream out = null;
        InputStream io = null;

        for (int i = 0; i < sheetList.size(); i++) {
            SheetInfoBean sheet = sheetList.get(i);

            XSSFSheet destSheet = destWorkbook.createSheet(sheet.getXls_work_sht_nm());

            if(i==0){ //기본 정보 입력

                String emailWorkTargId = templateCreationMap.get("email_work_targ_id") == null? UUID.randomUUID().toString() : templateCreationMap.get("email_work_targ_id") .toString();
                XSSFRow destRow = destSheet.createRow(0);
                CellStyle rowStyle = destRow.getRowStyle();
                DataFormat newDataFormat = destSheet.getWorkbook().createDataFormat();

                XSSFCell emailWorkTargIdNewCell = destRow.createCell(1); // new cell
                CellInfoBean emailWorkTargIdCell  = new CellInfoBean();
                emailWorkTargIdCell.setString_value(emailWorkTargId); //email_work_targ_id
                emailWorkTargIdCell.setCell_type(1);
                emailWorkTargIdCell.setHidden(true);
                emailWorkTargIdCell.setBorder_bottom_cd((short)0);
                emailWorkTargIdCell.setBorder_left_cd((short)0);
                emailWorkTargIdCell.setBorder_right_cd((short)0);
                emailWorkTargIdCell.setBorder_top_cd((short)0);
                emailWorkTargIdCell.setDataformat(newDataFormat.getFormat(";;;"));
                emailWorkTargIdCell.setCol_width(3);
                copyCell(emailWorkTargIdCell, emailWorkTargIdNewCell);



                String emailSndLogId = templateCreationMap.get("email_snd_log_id") == null? UUID.randomUUID().toString() : templateCreationMap.get("email_snd_log_id") .toString();
                XSSFCell emailSndLogIdNewCell = destRow.createCell(2); // new cell
                CellInfoBean emailSndLogIdCell  = new CellInfoBean();
                emailSndLogIdCell.setString_value(emailSndLogId); //email_work_targ_id
                emailSndLogIdCell.setCell_type(1);
                emailSndLogIdCell.setHidden(true);
                emailSndLogIdCell.setBorder_bottom_cd((short)0);
                emailSndLogIdCell.setBorder_left_cd((short)0);
                emailSndLogIdCell.setBorder_right_cd((short)0);
                emailSndLogIdCell.setBorder_top_cd((short)0);
                emailSndLogIdCell.setDataformat(newDataFormat.getFormat(";;;"));
                emailSndLogIdCell.setCol_width(3);
                copyCell(emailSndLogIdCell, emailSndLogIdNewCell);

            }

            //sheet별로 데이터를 다르게 저장해놨음.
            Map<String,Map<String,Map<String,Object>>> dataListRow = dataListSheet.get(sheet.getXls_work_sht());
            createExcelForMapListDataCreate(sheet.getRowList(), destSheet,dataListRow,headersList);
        }

    }


    //Template 만들때 사용하는 method
    public static Map<String,Object> createExcelByDataSetup(List<SheetInfoBean> sheetList, XSSFWorkbook destWorkbook,Map<String, Object> dataMapForSheet ,List<Map<String,Object>> headersList) {
        int maxColumnNum = 0;
        Map<String,Object> resultMap = new HashMap<String, Object>();

        createExcelByListDataSetup(sheetList,destWorkbook,dataMapForSheet,headersList);

        return resultMap;
    }


    //이메일 업무관리에서 메일 발송시 사용하는 Method
    public Map<String,Object> createExcelTemplateToDataSetup(List<SheetInfoBean> sheetList, Map<String,Map<String,Map<String,Map<String,Object>>>> dataListSheet ,List<Map<String,Object>> headersList,String fileNm) {
        int maxColumnNum = 0;
        Map<String,Object> resultMap = new HashMap<String, Object>();

        OutputStream out = null;
        Workbook destinationWorkBook = null;
        XSSFSheet destination = null;
        File file = null;
        InputStream io = null;

        try{
            //temp file
            file = new File(fileUploadPath  + UUID.randomUUID() + ".xlsx");

            destinationWorkBook = new XSSFWorkbook();

            out = new FileOutputStream(file);


            //우선 template를 만든다.
            resultMap = createExcelDataTemplateFirst(sheetList,dataListSheet,headersList);

            /**
             * resultMap
             * - email_work_targ_id
             * - email_snd_log_id
             * - tmp_form_file  ( template grp_cd )
             *
             */

            //template -> file read & sheet bean 객체로 변환
            List<SheetInfoBean> tmpSheetList = excelFileInsertProc(resultMap,sheetList);

            //data만 밀어넣는다.
            createExcelDataSetup(tmpSheetList,(XSSFWorkbook)destinationWorkBook,dataListSheet,headersList,resultMap);

            destinationWorkBook.write(out);

            io = FileUtils.openInputStream(file);

            String grpCd = UUID.randomUUID().toString();

            SimpleMultipartFileItem fileItem = new SimpleMultipartFileItem(UUID.randomUUID().toString(),grpCd, fileNm, FilenameUtils.getExtension(fileNm), file.length(), null, null, file);

            fileItem.setMultipartFile(new MockMultipartFile(fileNm, io));

            fileService.create(fileItem);

            resultMap.put("grp_cd",grpCd);
        } catch (IOException ioe) {
            LOG.error(ioe.getMessage());
        } catch (Exception e) {
            LOG.error(e.getMessage());
        } finally {
            try {
                if (destinationWorkBook != null) destinationWorkBook.close();
                if (out != null) out.close();
            } catch (Exception e) {
                LOG.error(e.getMessage());
            }
        }


        return resultMap;
    }


    /**
     * 메일 회신으로 넘어온 excel에 대해 첫번째 row / 첫번째 cell & 두번째 cell의 값을 찾아오는 메소드 ( 해당 값들을 이용하여, copy 처리 & 비교 처리 )
     * @param param
     * @param getSheetList
     * @return
     */
    public List<SheetInfoBean> excelFileInsertProc(Map<String,Object> param,List<SheetInfoBean> getSheetList){

        // SheetInfoBean List new
        List<SheetInfoBean> sheetList = new ArrayList<SheetInfoBean>();

        try{

            //미리 만들어둔 temp file grp_cd
            String attNo = param.get("tmp_form_file") == null? "" : param.get("tmp_form_file").toString();


            FileGroup group = fileService.findGroup(attNo);
            if(group == null || group.getSize() == 0){
                throw new FileNotFoundException("첨부된 파일이 없는 메일 입니다!");
            }

            FileItem excelFileItem = null;
            for(FileItem fileItem :group.getItems()){
                if("xlsx".equals(fileItem.getExtension())){
                    excelFileItem = fileItem;
                    break;
                }else if(fileItem.getName().indexOf("xlsx") > -1){
                    excelFileItem = fileItem;
                    break;
                }

            }

            if(excelFileItem == null){
                throw new FileNotFoundException("엑셀파일이 없습니다");
            }

            try {
                excelFileItem = fileService.findDownloadItem(excelFileItem.getId());
            } catch (Exception e1) {
                throw new FileNotFoundException("파일을 가져오는 중 오류발생!");
            }

            if(excelFileItem.getFile().exists()){


                // 화면단에서 정의한 attachment의 file grp_cd를 가지고 취득하여 excel을 가져온다.
                Workbook sourceWorkBook = new XSSFWorkbook(OPCPackage.open(excelFileItem.getFile().getPath()));

                //해당 엑셀 파일 내에 존재하는 Sheet의 갯수를 가져온다.
                int sheetCnt = sourceWorkBook.getNumberOfSheets();

                // SheetCnt 갯수만큼 시트별 데이터를 bean에 담는다.
                for(int i = 0; i < sheetCnt; i++) {

                    SheetInfoBean sheetInfo = new SheetInfoBean();

                    // Excel에 대한 데이터를 가져온다.
                    XSSFSheet source = ((XSSFWorkbook) sourceWorkBook).getSheetAt(i);

                    for(SheetInfoBean getSheetInfo : getSheetList){

                        String sourceSheetName = source.getSheetName() == null? "" : source.getSheetName();
                        String getSheetInfoSheetName =  getSheetInfo.getXls_work_sht_nm() == null? "" :  getSheetInfo.getXls_work_sht_nm();

                        //SheetName은 독립적이기에 비교하여, 처리 가능. SheetInfo를 여기서 가져옴.
                        if(sourceSheetName.equals(getSheetInfoSheetName)) sheetInfo = getSheetInfo; break;
                    }

                    String emailWorkId = sheetInfo.getEmail_work_id();
                    String xlsWorkSht = sheetInfo.getXls_work_sht();

                    //Excel에 ROW / CELL 정보를 취합한다.
                    List<RowInfoBean> sheetRow = excelReaderUtil.readExcel(source,emailWorkId,xlsWorkSht);

                    //EXCEL에서 읽어온 정보를 기준으로 ROW LIST를 SET 한다.
                    sheetInfo.setRowList(sheetRow);

                    //ArrayList add
                    sheetList.add(sheetInfo);
                }
            }
        }catch (RuntimeException rune){
            LOG.error(rune.getMessage());
        }catch (Exception e){
            LOG.error(e.getMessage());
        }

        return sheetList;

    }


    /**
     * LIST의 갯수를 체크하면서, LIST 밑에 있는 단일 객체도 치환이 정상적으로 가능하도록 하여, 치환되도록 하는 Method
     * @param rowList
     * @param destination
     * @param dataMapForRow
     * @param headersList
     */
    public static void createExcelForMapListDataCreate(List<RowInfoBean> rowList, XSSFSheet destination , Map<String,Map<String,Map<String,Object>>> dataMapForRow , List<Map<String,Object>> headersList ) {
        int maxColumnNum = 0;
        int headerNextRow = 0;
        int dumyRowCnt = 0;

        for (int i = 0; i < rowList.size(); i++) {
            RowInfoBean sheetRow = rowList.get(i);
            int rowCreateNum = sheetRow.getRow_no();

            int replaceType = checkCellListAndValue(sheetRow);

            if(replaceType == ExcelCreateUtil.REPLACE_TYPE_LIST){

                String getRowNo = Integer.toString(sheetRow.getRow_no());
                Map<String,Map<String,Object>> dataList = dataMapForRow.get(getRowNo);

                int lastDataList = dataList.size();

                for(int a = 0; a < lastDataList; a++){
                    Map<String,Object> data = dataList.get(Integer.toString(rowCreateNum));
                    //data list 를 기준으로 copy 처리한다.
                    XSSFRow destRow = destination.createRow(rowCreateNum);
                    copyRow(sheetRow, destRow);
                    dataSetRowForMapList( destRow , data);
                    dumyRowCnt++;
                }

            }else if(replaceType == ExcelCreateUtil.REPLACE_TYPE_VALUE){

                int checkDataCnt = 0;

                if(dumyRowCnt > 0){
                    checkDataCnt = rowCreateNum - dumyRowCnt;
                }else{
                    checkDataCnt = rowCreateNum;
                }

                Map<String,Map<String,Object>> dataList = dataMapForRow.get(Integer.toString(checkDataCnt));

                Map<String,Object> data = dataList.get(Integer.toString(checkDataCnt));

                XSSFRow destRow = destination.createRow(rowCreateNum);
                copyRow(sheetRow, destRow);

                //data list 를 기준으로 copy 처리한다.
                dataSetRowForMapValue( destination.getRow(rowCreateNum) , data);

            }else{
                XSSFRow destRow = destination.createRow(rowCreateNum);
                copyRow(sheetRow, destRow);
            }


            /*if(headersList.size() > 0 && i == 0){
                XSSFRow destRow = destination.createRow(sheetRow.getRowNum());
                setHeadersRow(sheetRow, destRow,headersList);
                headerNextRow = sheetRow.getRowNum() + 1;*/

        }
        for (int i = 0; i < rowList.size(); i++) {
            RowInfoBean sheetRow = rowList.get(i);
            destination.autoSizeColumn(i);
            //destination.setColumnWidth(i, sheetRow.getColumn_width());
        }
    }



    // DATA를 받아와 해당 데이터만큼의 VALUE의 SIZE를 계산하여, 미리 TEMPLATE를 생성
    public static void createExcelForMapListTemplateCreate(List<RowInfoBean> rowList, XSSFSheet destination , Map<String,Map<String,Map<String,Object>>> dataMapForRow , List<Map<String,Object>> headersList ) {
        int maxColumnNum = 0;
        int headerNextRow = 0;

        int lastDataRow = 0;
        for (int i = 0; i < rowList.size(); i++) {
            RowInfoBean sheetRow = rowList.get(i);
            int rowNo = sheetRow.getRow_no();
            int rowCreateNum = 0;

            if(rowNo == 0) continue;

            String getRowNo = Integer.toString(sheetRow.getRow_no());
            Map<String,Map<String,Object>> dataList = dataMapForRow.get(getRowNo);

            int replaceType = checkCellListAndValue(sheetRow);

            if(replaceType == ExcelCreateUtil.REPLACE_TYPE_LIST){

                if (rowCreateNum == 0)  rowCreateNum = rowNo;

                int lastDataList = dataList.size();

                for(int b = 0; b < lastDataList; b++){
                    XSSFRow destRow = destination.createRow(rowCreateNum);
                    copyRow(sheetRow, destRow);
                    rowCreateNum++;
                    lastDataRow = rowCreateNum;
                }

            }else if(replaceType == ExcelCreateUtil.REPLACE_TYPE_VALUE){

                if (rowCreateNum == 0)  rowCreateNum = rowNo;
                int lastDataList = dataList.size();
                if(lastDataRow != 0) {
                    XSSFRow destRow = destination.createRow(rowCreateNum+lastDataList);
                    copyRow(sheetRow, destRow);
                }else{
                    XSSFRow destRow = destination.createRow(rowCreateNum);
                    copyRow(sheetRow, destRow);
                }
            }else{
                if (rowCreateNum == 0)  rowCreateNum = rowNo;
                XSSFRow destRow = destination.createRow(rowCreateNum);
                copyRow(sheetRow, destRow);
            }
        }
        for (int i = 0; i < rowList.size(); i++) {
            RowInfoBean sheetRow = rowList.get(i);
            destination.autoSizeColumn(i);
            //destination.setColumnWidth(i, sheetRow.getColumn_width());
        }
    }


    // 업무에서 map을 던지면 template 생성 및 data 매칭하는 중간 Method
    public Map<String,Object> createEmailWorkProc(List<SheetInfoBean> sheetList, Map<String,Object> dataMap ,List<Map<String,Object>> headersList,String fileNm) {
        int maxColumnNum = 0;
        Map<String,Object> resultMap = new HashMap<String, Object>();

        OutputStream out = null;
        Workbook destinationWorkBook = null;
        XSSFSheet destination = null;
        File file = null;
        InputStream io = null;

        try{
            //temp file
            file = new File(fileUploadPath  + UUID.randomUUID() + ".xlsx");

            destinationWorkBook = new XSSFWorkbook();

            out = new FileOutputStream(file);


            //우선 template를 만든다.
            resultMap = createEmailWorkExcelTemplate(sheetList,dataMap,headersList);

            /**
             * resultMap
             * - email_work_targ_id
             * - email_snd_log_id
             * - tmp_form_file  ( template grp_cd )
             *
             */

            //template -> file read & sheet bean 객체로 변환
            List<SheetInfoBean> tmpSheetList = excelFileInsertProc(resultMap,sheetList);

            //data만 밀어넣는다.
            createEmailWorkExcelDataSetup(tmpSheetList,(XSSFWorkbook)destinationWorkBook,dataMap,headersList,resultMap);

            destinationWorkBook.write(out);

            io = FileUtils.openInputStream(file);

            String grpCd = UUID.randomUUID().toString();

            SimpleMultipartFileItem fileItem = new SimpleMultipartFileItem(UUID.randomUUID().toString(),grpCd, fileNm, FilenameUtils.getExtension(fileNm), file.length(), null, null, file);

            fileItem.setMultipartFile(new MockMultipartFile(fileNm, io));

            fileService.create(fileItem);

            resultMap.put("grp_cd",grpCd);
        } catch (IOException ioe) {
            LOG.error(ioe.getMessage());
        } catch (Exception e) {
            LOG.error(e.getMessage());
        } finally {
            try {
                if (destinationWorkBook != null) destinationWorkBook.close();
                if (out != null) out.close();
            } catch (Exception e) {
                LOG.error(e.getMessage());
            }
        }


        return resultMap;
    }


    /**
     * 업무에서 이메일 견적을 생성할 시 추후 회신 시에 넘어오는 excel 파일이 정상적인지 판단하기 위함 및 cell index 위치 및 row의 치환을 위한 여러가지 방편으로 data map에 있는 list 및 각 객체를 위하여 발송전 template를 생성
     * 해당 파일은 tmp_form_file 로 저장되며, 회신 온 메일과 비교 분석하여, result map으로 나타나게 합니다.
     *
     * @param sheetList
     * @param dataListSheet
     * @param headersList
     * @return
     */
    public Map<String,Object> createEmailWorkExcelTemplate(List<SheetInfoBean> sheetList, Map<String,Object> dataMap ,List<Map<String,Object>> headersList){

        Map<String,Object> resultMap = new HashMap<String, Object>();
        OutputStream out = null;
        InputStream io = null;
        XSSFWorkbook tempWorkBook = null;
        File file = null;
        try{

            String fileNm = UUID.randomUUID().toString() + ".xlsx";

            file = new File(fileUploadPath + FilenameUtils.getExtension(fileNm));

            tempWorkBook = new XSSFWorkbook();

            out = new FileOutputStream(file);

            for (int i = 0; i < sheetList.size(); i++) {

                SheetInfoBean sheet = sheetList.get(i);


                XSSFSheet destSheet = tempWorkBook.createSheet(sheet.getXls_work_sht_nm());

                if(i==0){ //기본 정보 입력

                    String emailWorkTargId = UUID.randomUUID().toString();
                    XSSFRow destRow = destSheet.createRow(0);
                    CellStyle rowStyle = destRow.getRowStyle();
                    DataFormat newDataFormat = destSheet.getWorkbook().createDataFormat();

                    XSSFCell emailWorkTargIdNewCell = destRow.createCell(1); // new cell
                    CellInfoBean emailWorkTargIdCell  = new CellInfoBean();
                    emailWorkTargIdCell.setString_value(emailWorkTargId); //email_work_targ_id
                    emailWorkTargIdCell.setCell_type(1);
                    emailWorkTargIdCell.setHidden(true);
                    emailWorkTargIdCell.setBorder_bottom_cd((short)0);
                    emailWorkTargIdCell.setBorder_left_cd((short)0);
                    emailWorkTargIdCell.setBorder_right_cd((short)0);
                    emailWorkTargIdCell.setBorder_top_cd((short)0);
                    emailWorkTargIdCell.setDataformat(newDataFormat.getFormat(";;;"));
                    emailWorkTargIdCell.setCol_width(3);
                    copyCell(emailWorkTargIdCell, emailWorkTargIdNewCell);

                    String emailSndLogId = UUID.randomUUID().toString();
                    XSSFCell emailSndLogIdNewCell = destRow.createCell(2); // new cell
                    CellInfoBean emailSndLogIdCell  = new CellInfoBean();
                    emailSndLogIdCell.setString_value(emailSndLogId); //email_work_targ_id
                    emailSndLogIdCell.setCell_type(1);
                    emailSndLogIdCell.setHidden(true);
                    emailSndLogIdCell.setBorder_bottom_cd((short)0);
                    emailSndLogIdCell.setBorder_left_cd((short)0);
                    emailSndLogIdCell.setBorder_right_cd((short)0);
                    emailSndLogIdCell.setBorder_top_cd((short)0);
                    emailSndLogIdCell.setDataformat(newDataFormat.getFormat(";;;"));
                    emailSndLogIdCell.setCol_width(3);
                    copyCell(emailSndLogIdCell, emailSndLogIdNewCell);

                    resultMap.put("email_work_targ_id",emailWorkTargId);
                    resultMap.put("email_snd_log_id",emailSndLogId);
                }

                createEmailWorkExcelForMapTemplateCreate(sheet.getRowList(), destSheet,dataMap,headersList);
            }


            tempWorkBook.write(out);

            io = FileUtils.openInputStream(file);

            String grpCd = UUID.randomUUID().toString();

            SimpleMultipartFileItem fileItem = new SimpleMultipartFileItem(UUID.randomUUID().toString(),grpCd, fileNm, FilenameUtils.getExtension(fileNm), file.length(), null, null, file);

            fileItem.setMultipartFile(new MockMultipartFile(fileNm, io));

            fileService.create(fileItem);

            resultMap.put("tmp_form_file",grpCd);  // template용 file
        }catch (Exception e){
            LOG.error(e.getMessage());
        }



        return resultMap;
    }

    // DATA를 받아와 해당 데이터만큼의 VALUE의 SIZE를 계산하여, 미리 TEMPLATE를 생성
    public static void createEmailWorkExcelForMapTemplateCreate(List<RowInfoBean> rowList, XSSFSheet destination , Map<String,Object> dataMap , List<Map<String,Object>> headersList ) {
        int maxColumnNum = 0;
        int headerNextRow = 0;

        int lastDataRow = 0;

        Map<String,Object> listDataMap = new HashMap<String, Object>();

        for (int i = 0; i < rowList.size(); i++) {
            RowInfoBean sheetRow = rowList.get(i);
            int rowNo = sheetRow.getRow_no();
            int rowCreateNum = 0;

            if(rowNo == 0) continue;

            Map<String,Object> checkValueMap = checkCellListAndValueRetrunMapType(sheetRow);
            int replaceType = checkValueMap.get("checkReplaceType") == null? 0 : Integer.parseInt(checkValueMap.get("checkReplaceType").toString()); //value의 list or object type
            String getValue = checkValueMap.get("getValue") == null? "" : checkValueMap.get("getValue").toString();  // data의 map을 체크하는 value



            if(replaceType == ExcelCreateUtil.REPLACE_TYPE_LIST){
                List<Map<String,Object>> dataList = (List<Map<String, Object>>) dataMap.get(getValue);

                if (rowCreateNum == 0)  rowCreateNum = rowNo;

                int lastDataList = dataList.size();

                for(int b = 0; b < lastDataList; b++){
                    XSSFRow destRow = destination.createRow(rowCreateNum);
                    copyRow(sheetRow, destRow);
                    listDataMap.put(Integer.toString(rowCreateNum) , dataList.get(b));
                    rowCreateNum++;
                }
                lastDataRow = lastDataList;

            }else if(replaceType == ExcelCreateUtil.REPLACE_TYPE_VALUE){
                Map<String,Object> dataList = (Map<String,Object>)dataMap.get(getValue);

                if (rowCreateNum == 0)  rowCreateNum = rowNo;
                if(lastDataRow != 0) {
                    XSSFRow destRow = destination.createRow(rowCreateNum+lastDataRow);
                    copyRow(sheetRow, destRow);
                }else{
                    XSSFRow destRow = destination.createRow(rowCreateNum);
                    copyRow(sheetRow, destRow);
                }
            }else{
                if (rowCreateNum == 0)  rowCreateNum = rowNo;
                XSSFRow destRow = destination.createRow(rowCreateNum+lastDataRow);
                copyRow(sheetRow, destRow);
            }
        }

        dataMap.put("listDataMap",listDataMap);
        for (int i = 0; i < rowList.size(); i++) {
            RowInfoBean sheetRow = rowList.get(i);
            destination.autoSizeColumn(i);
            //destination.setColumnWidth(i, sheetRow.getColumn_width());
        }
    }

    /**
     * Cell의 String Value내에 객체를 표기하는 형태가 존재하는지 찾고, 분류값을 지정한다. ( 해당 Method는 Row 내에 1개의 구분값(리스트,단일객체)가 있다는 전재하이다.)
     * -- List 객체의 경우  $[ ]
     * -- 단일 객체의 경우  ${ }
     * @param srcRow
     * @return
     */
    private static Map<String,Object> checkCellListAndValueRetrunMapType(  RowInfoBean srcRow) {
        boolean checkReplaceValue = false;
        boolean checkReplaceList = false;

        // 현재까지 기준으론 1개의 row에 단일 또는 리스트로만 정의하도록 한다.
        int checkReplaceType = 0;
        String getValue = "";
        Map<String,Object> resultMap = new HashMap<String, Object>();


        List<CellInfoBean> rowList = srcRow.getCellList();

        for (int i = 0; i < rowList.size(); i++) {
            CellInfoBean cellInfoBean = rowList.get(i);
            getValue = cellInfoBean.getString_value() == null? "" : cellInfoBean.getString_value();

            if(!StringUtils.isEmpty(getValue)){
                if (getValue.startsWith(ExcelCreateUtil.STARTEXPRESSIONTOKEN) && getValue.endsWith(ExcelCreateUtil.ENDEXPRESSIONTOKEN)) {
                    checkReplaceValue = true;
                    getValue = getValue.replace(ExcelCreateUtil.STARTEXPRESSIONTOKEN,"");
                    getValue = getValue.replace(ExcelCreateUtil.ENDEXPRESSIONTOKEN,"");

                    if(!StringUtils.isEmpty(getValue)){
                        String dataGetKey = "";
                        if(getValue.indexOf(".") > -1){ // ${ data.value } 형식의 구조라면, 아래와 같이.
                            String[] splitValue = getValue.split("\\.");
                            getValue = splitValue[0];
                        }
                    }
                    break;
                }else if(getValue.startsWith(ExcelCreateUtil.STARTFORMULATOKEN) && getValue.endsWith(ExcelCreateUtil.ENDFORMULATOKEN)){
                    checkReplaceList = true;
                    getValue = getValue.replace(ExcelCreateUtil.STARTFORMULATOKEN,"");
                    getValue = getValue.replace(ExcelCreateUtil.ENDFORMULATOKEN,"");

                    if(!StringUtils.isEmpty(getValue)){
                        String dataGetKey = "";
                        if(getValue.indexOf(".") > -1){ // ${ data.value } 형식의 구조라면, 아래와 같이.
                            String[] splitValue = getValue.split("\\.");
                            getValue = splitValue[0];
                        }
                    }
                    break;
                }
            }
        }


        if(checkReplaceList){  // 리스트
            checkReplaceType = ExcelCreateUtil.REPLACE_TYPE_LIST;
        }else if(checkReplaceValue){  // 단일
            checkReplaceType = ExcelCreateUtil.REPLACE_TYPE_VALUE;
        }else{
            checkReplaceType = ExcelCreateUtil.REPLACE_TYPE_NONE;
        }

        resultMap.put("checkReplaceType",checkReplaceType);
        resultMap.put("getValue",getValue);
        return resultMap;
    }


    /**
     * template 로 만들어둔 excel 파일의 변수명에 맞춰 넘어온 data map과 맞춰 치환 시키는 method
     * @param sheetList
     * @param destWorkbook
     * @param dataListSheet
     * @param headersList
     * @param templateCreationMap
     */
    public static  void createEmailWorkExcelDataSetup(List<SheetInfoBean> sheetList, XSSFWorkbook destWorkbook,Map<String,Object> dataMap ,List<Map<String,Object>> headersList , Map<String,Object> templateCreationMap){

        Map<String,Object> resultMap = new HashMap<String, Object>();
        OutputStream out = null;
        InputStream io = null;

        for (int i = 0; i < sheetList.size(); i++) {
            SheetInfoBean sheet = sheetList.get(i);

            XSSFSheet destSheet = destWorkbook.createSheet(sheet.getXls_work_sht_nm());

            if(i==0){ //기본 정보 입력

                String emailWorkTargId = templateCreationMap.get("email_work_targ_id") == null? UUID.randomUUID().toString() : templateCreationMap.get("email_work_targ_id") .toString();
                XSSFRow destRow = destSheet.createRow(0);
                CellStyle rowStyle = destRow.getRowStyle();
                DataFormat newDataFormat = destSheet.getWorkbook().createDataFormat();

                XSSFCell emailWorkTargIdNewCell = destRow.createCell(1); // new cell
                CellInfoBean emailWorkTargIdCell  = new CellInfoBean();
                emailWorkTargIdCell.setString_value(emailWorkTargId); //email_work_targ_id
                emailWorkTargIdCell.setCell_type(1);
                emailWorkTargIdCell.setHidden(true);
                emailWorkTargIdCell.setBorder_bottom_cd((short)0);
                emailWorkTargIdCell.setBorder_left_cd((short)0);
                emailWorkTargIdCell.setBorder_right_cd((short)0);
                emailWorkTargIdCell.setBorder_top_cd((short)0);
                emailWorkTargIdCell.setDataformat(newDataFormat.getFormat(";;;"));
                emailWorkTargIdCell.setCol_width(3);
                copyCell(emailWorkTargIdCell, emailWorkTargIdNewCell);

                String emailSndLogId = templateCreationMap.get("email_snd_log_id") == null? UUID.randomUUID().toString() : templateCreationMap.get("email_snd_log_id") .toString();
                XSSFCell emailSndLogIdNewCell = destRow.createCell(2); // new cell
                CellInfoBean emailSndLogIdCell  = new CellInfoBean();
                emailSndLogIdCell.setString_value(emailSndLogId); //email_work_targ_id
                emailSndLogIdCell.setCell_type(1);
                emailSndLogIdCell.setHidden(true);
                emailSndLogIdCell.setBorder_bottom_cd((short)0);
                emailSndLogIdCell.setBorder_left_cd((short)0);
                emailSndLogIdCell.setBorder_right_cd((short)0);
                emailSndLogIdCell.setBorder_top_cd((short)0);
                emailSndLogIdCell.setDataformat(newDataFormat.getFormat(";;;"));
                emailSndLogIdCell.setCol_width(3);
                copyCell(emailSndLogIdCell, emailSndLogIdNewCell);

            }

            createEmailWorkExcelForDataMapMapping(sheet.getRowList(), destSheet,dataMap,headersList);
        }

    }

    /**
     * LIST의 갯수를 체크하면서, LIST 밑에 있는 단일 객체도 치환이 정상적으로 가능하도록 하여, 치환되도록 하는 Method
     * @param rowList
     * @param destination
     * @param dataMapForRow
     * @param headersList
     */
    public static void createEmailWorkExcelForDataMapMapping(List<RowInfoBean> rowList, XSSFSheet destination , Map<String,Object> dataMap , List<Map<String,Object>> headersList ) {
        int maxColumnNum = 0;
        int headerNextRow = 0;
        int dumyRowCnt = 0;

        for (int i = 0; i < rowList.size(); i++) {
            RowInfoBean sheetRow = rowList.get(i);
            int rowCreateNum = sheetRow.getRow_no();

            Map<String,Object> checkValueMap = checkCellListAndValueRetrunMapType(sheetRow);
            int replaceType = checkValueMap.get("checkReplaceType") == null? 0 : Integer.parseInt(checkValueMap.get("checkReplaceType").toString()); //value의 list or object type
            String getValue = checkValueMap.get("getValue") == null? "" : checkValueMap.get("getValue").toString();  // data의 map을 체크하는 value


            if(replaceType == ExcelCreateUtil.REPLACE_TYPE_LIST){
                String rowNo = Integer.toString(sheetRow.getRow_no());

                Map<String,Object> listDataMap = (Map<String,Object>)dataMap.get("listDataMap");

                //data list 를 기준으로 copy 처리한다.
                XSSFRow destRow = destination.createRow(rowCreateNum);
                copyRow(sheetRow, destRow);
                dataSetRowForDataMapValue( destRow ,   (Map<String,Object>)listDataMap.get(rowNo));
            }else if(replaceType == ExcelCreateUtil.REPLACE_TYPE_VALUE){

                int checkDataCnt = 0;
                Map<String,Object> dataList = (Map<String,Object>)dataMap.get(getValue);
                XSSFRow destRow = destination.createRow(rowCreateNum);
                copyRow(sheetRow, destRow);

                //data list 를 기준으로 copy 처리한다.
                dataSetRowForDataMapValue( destination.getRow(rowCreateNum) , dataList);

            }else{
                XSSFRow destRow = destination.createRow(rowCreateNum);
                copyRow(sheetRow, destRow);
            }


            /*if(headersList.size() > 0 && i == 0){
                XSSFRow destRow = destination.createRow(sheetRow.getRowNum());
                setHeadersRow(sheetRow, destRow,headersList);
                headerNextRow = sheetRow.getRowNum() + 1;*/

        }
        for (int i = 0; i < rowList.size(); i++) {
            RowInfoBean sheetRow = rowList.get(i);
            destination.autoSizeColumn(i);
            //destination.setColumnWidth(i, sheetRow.getColumn_width());
        }
    }


    private static void dataSetRowForDataMapValue(  XSSFRow destRow , Map<String,Object> data) {
        //Set<CellRangeAddressWrapper> mergedRegions = new TreeSet<CellRangeAddressWrapper>();

        for (int j = 0; j <= destRow.getLastCellNum(); j++) {
            XSSFCell getCell = destRow.getCell(j); // ancienne cell
            if(getCell == null) continue;

            String excelStringValue = getCell.getStringCellValue() == null? "" : getCell.getStringCellValue();
            String getValue = "";

            if (excelStringValue.startsWith(ExcelCreateUtil.STARTEXPRESSIONTOKEN) && excelStringValue.endsWith(ExcelCreateUtil.ENDEXPRESSIONTOKEN)) {
                String dataGetKey = "";
                excelStringValue = excelStringValue.replace(ExcelCreateUtil.STARTEXPRESSIONTOKEN,"");
                excelStringValue = excelStringValue.replace(ExcelCreateUtil.ENDEXPRESSIONTOKEN,"");

                if(excelStringValue.indexOf(".") > -1){ // ${ data.value } 형식의 구조라면, 아래와 같이.

                    String[] splitValue = excelStringValue.split("\\.");
                    dataGetKey = splitValue[0];
                    excelStringValue = splitValue[1];
                    getValue = data.get(excelStringValue) == null? "" : data.get(excelStringValue).toString();
                }else{ // ${ value } 형식의 구조라면, 아래와 같이.
                    getValue = data.get(excelStringValue) == null? "" : data.get(excelStringValue).toString();
                }
            } else if (excelStringValue.startsWith(ExcelCreateUtil.STARTFORMULATOKEN) && excelStringValue.endsWith(ExcelCreateUtil.ENDFORMULATOKEN)) {
                String dataGetKey = "";
                excelStringValue = excelStringValue.replace(ExcelCreateUtil.STARTFORMULATOKEN,"");
                excelStringValue = excelStringValue.replace(ExcelCreateUtil.ENDFORMULATOKEN,"");

                if(excelStringValue.indexOf(".") > -1){ // ${ data.value } 형식의 구조라면, 아래와 같이.
                    String[] splitValue = excelStringValue.split("\\.");
                    dataGetKey = splitValue[0];
                    excelStringValue = splitValue[1];
                    getValue = data.get(excelStringValue) == null? "" : data.get(excelStringValue).toString();
                }else{ // ${ value } 형식의 구조라면, 아래와 같이.
                    getValue = data.get(excelStringValue) == null? "" : data.get(excelStringValue).toString();
                }
            }else{
                if(StringUtils.isEmpty(getValue)){
                    getValue = getCell.getStringCellValue() == null? "" : getCell.getStringCellValue();
                }
            }

            getCell.setCellValue(getValue);
        }
    }

    public static void main(String[] args) throws IOException, InvalidFormatException {

        Workbook sourceWorkBook = new XSSFWorkbook(OPCPackage.open("C:\\solution\\testfile\\test_1.xlsx"));
        Workbook destinationWorkBook = new XSSFWorkbook(OPCPackage.open("C:\\solution\\testfile\\test_2.xlsx"));
        OutputStream out = new FileOutputStream(new File("C:\\solution\\testfile\\testFile_out.xlsx"));

      /*  wb_1.setSheetName(0, "Actual");
        wb_1.createSheet("Last Week");*/

        /*Get sheets from the temp file*/
        XSSFSheet destination = ((XSSFWorkbook) destinationWorkBook).getSheetAt(0);
        XSSFSheet source = ((XSSFWorkbook) sourceWorkBook).getSheetAt(0);

        //copySheet(source, destination);

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
