package smartsuite.app.common.excel;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.filefilter.TrueFileFilter;
import org.apache.commons.lang.StringUtils;
import org.apache.ibatis.session.SqlSession;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.BeanUtils;
import org.springframework.stereotype.Service;
import smartsuite.app.common.cert.util.EdocStringUtil;
import smartsuite.app.common.excel.bean.CellInfoBean;
import smartsuite.app.common.excel.bean.RowInfoBean;
import smartsuite.app.common.excel.bean.SheetInfoBean;
import smartsuite.upload.core.entity.FileGroup;
import smartsuite.upload.core.entity.FileItem;
import smartsuite.upload.core.service.FileService;

import javax.inject.Inject;
import javax.servlet.ServletOutputStream;
import java.io.*;
import java.sql.*;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.Date;

@Service
@SuppressWarnings("unused")
public class ExcelReaderUtil {

    static final Logger LOG = LoggerFactory.getLogger(ExcelReaderUtil.class);


    @Inject
    FileService fileService;

    @Inject
    private SqlSession sqlSession;


    // value 로만 단일 값에 대해서 처리하기로함
    private static final String STARTEXPRESSIONTOKEN = "${";
    private static final String ENDEXPRESSIONTOKEN = "}";

    // list 처리할때 사용 하기로 함
    private static final String STARTFORMULATOKEN = "$[";
    private static final String ENDFORMULATOKEN = "]";



    public static List<RowInfoBean> readExcel(XSSFSheet source , String emailWorkId,String xlsWorkSht) {

        List<RowInfoBean> rowInfoBeanList = new ArrayList<RowInfoBean>();
        for (int i = source.getFirstRowNum(); i <= source.getLastRowNum(); i++) {
            XSSFRow srcRow = source.getRow(i);
            RowInfoBean sheetRow = new RowInfoBean();
            String rowKey = UUID.randomUUID().toString();
            sheetRow.setRow_no(i);
            if (srcRow != null) {
                sheetRow.setRow_id(rowKey);
                sheetRow.setEmail_work_id(emailWorkId);
                sheetRow.setXls_work_sht(xlsWorkSht);
                sheetRow.setCellList(readRow(source,  srcRow ,emailWorkId , rowKey));
                //sheetRow.setColumnIndex(i);
                rowInfoBeanList.add(sheetRow);
            }
        }
        return rowInfoBeanList;
    }


    private static List<CellInfoBean> readRow(XSSFSheet srcSheet, XSSFRow srcRow , String emailWorkId , String rowKey) {
        short dh = srcSheet.getDefaultRowHeight();
        if (srcRow.getHeight() != dh) { //NOPMD
            //
            //destRow.setHeight(srcRow.getHeight());
        }

        List<CellInfoBean> cellInfoBeanList = new ArrayList<CellInfoBean>();


        int j = srcRow.getFirstCellNum();
        if (j < 0) {
            j = 0;
        }

        for (; j <= srcRow.getLastCellNum(); j++) {
            XSSFCell oldCell = srcRow.getCell(j); // ancienne cell
             if (oldCell != null) {
                 cellInfoBeanList.add(readCell(oldCell,j,emailWorkId,rowKey));
            }
        }

        return cellInfoBeanList;
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

    private static CellInfoBean readCell(Cell oldCell, int cellIndex, String emailWorkId , String rowKey) {


            XSSFFont oldFont = (XSSFFont) oldCell.getSheet().getWorkbook().getFontAt(oldCell.getCellStyle().getFontIndex());
            CellInfoBean cellInfoBean = new CellInfoBean();
            cellInfoBean.setCell_id(UUID.randomUUID().toString());
            cellInfoBean.setSheet_id(emailWorkId);
            cellInfoBean.setRow_id(rowKey);
            cellInfoBean.setBold(oldFont.getBold());
            cellInfoBean.setFont_height(oldFont.getFontHeight());
            cellInfoBean.setFont_nm(oldFont.getFontName());
            cellInfoBean.setItalic(oldFont.getItalic());
            cellInfoBean.setStrikeout(oldFont.getStrikeout());
            cellInfoBean.setType_offset(oldFont.getTypeOffset());
            cellInfoBean.setUnder_line(oldFont.getUnderline());
            cellInfoBean.setCharset(oldFont.getCharSet());
            cellInfoBean.setColor(getColorRGB(oldFont.getXSSFColor()));
            cellInfoBean.setDataformat(oldCell.getCellStyle().getDataFormat());
            cellInfoBean.setAlignment_cd(oldCell.getCellStyle().getAlignment());
            cellInfoBean.setHidden(oldCell.getCellStyle().getHidden());
            cellInfoBean.setLocked(oldCell.getCellStyle().getLocked());
            cellInfoBean.setWraptext(oldCell.getCellStyle().getWrapText());
            cellInfoBean.setBorder_bottom_cd(oldCell.getCellStyle().getBorderBottom());
            cellInfoBean.setBorder_left_cd(oldCell.getCellStyle().getBorderLeft());
            cellInfoBean.setBorder_right_cd(oldCell.getCellStyle().getBorderRight());
            cellInfoBean.setBorder_top_cd(oldCell.getCellStyle().getBorderTop());
            cellInfoBean.setBottom_border_color(getBottomBorderColorRGB((XSSFCellStyle)oldCell.getCellStyle()));
            cellInfoBean.setFill_background_color(getFillBackgroundColorRGB((XSSFCellStyle)oldCell.getCellStyle()));
            cellInfoBean.setFill_foreground_color(getFillForegroundColorRGB((XSSFCellStyle)oldCell.getCellStyle()));
            cellInfoBean.setFill_pattern(oldCell.getCellStyle().getFillPattern());
            cellInfoBean.setIndention(oldCell.getCellStyle().getIndention());
            cellInfoBean.setLeft_border_color(getLeftBorderColorRGB((XSSFCellStyle)oldCell.getCellStyle()));
            cellInfoBean.setRight_border_color(getRightBorderColorRGB((XSSFCellStyle)oldCell.getCellStyle()));
            cellInfoBean.setRotation(oldCell.getCellStyle().getRotation());
            cellInfoBean.setTop_border_color(getTopBorderColorRGB((XSSFCellStyle)oldCell.getCellStyle()));
            cellInfoBean.setVertical_alignment_cd(oldCell.getCellStyle().getVerticalAlignment());
            cellInfoBean.setCell_index(cellIndex);
            int cellColumnIndex =  oldCell.getSheet().getColumnWidth(oldCell.getColumnIndex());
            cellInfoBean.setCol_width(cellColumnIndex);


            /**
         *      STRING: 1
         *      NUMERIC: 0
         *      BLANK: 4
         *      BOOLEAN: 5
         *      ERROR: 6
         *      FORMULA: 2
         *      DEFAULT : 0 -> 해당 코드는 NULL 처리 되어야함.
         * */

        switch (oldCell.getCellType()) {
                case 1:
                    cellInfoBean.setCell_type(1);
                    cellInfoBean.setString_value(oldCell.getStringCellValue());
                break;
                case 0:
                    if( DateUtil.isCellDateFormatted(oldCell)) {
                        Date date = oldCell.getDateCellValue();
                        String cellString = new SimpleDateFormat("yyyy-MM-dd", Locale.getDefault()).format(date);
                        cellInfoBean.setCell_type(1);
                        cellInfoBean.setString_value(cellString);
                    }else{
                        cellInfoBean.setCell_type(0);
                        cellInfoBean.setDouble_value(oldCell.getNumericCellValue());
                    }
                break;
                case 3:
                    cellInfoBean.setCell_type(3);
                break;
                case 4:
                    cellInfoBean.setCell_type(4);
                    cellInfoBean.setBoolean_value(oldCell.getBooleanCellValue());
                break;
                case 5:
                    cellInfoBean.setCell_type(5);
                    cellInfoBean.setError_value(oldCell.getErrorCellValue());
                break;
                case 2:
                    cellInfoBean.setCell_type(2);
                    cellInfoBean.setFormula_value(oldCell.getCellFormula());
                break;
                default:
                break;
        }

        return cellInfoBean;
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

    private static XSSFColor getColor(XSSFColor c) {
        if (c == null) {
            return null;
        }else if (c instanceof XSSFColor) {
            XSSFColor xc = c;
            byte[] data = null;
            if (xc.getTint() != 0.0) {
                data = getRgbWithTint(xc);
                byte[] argb = xc.getARGB();
            } else {
                data = xc.getARGB();
            }
            if (data == null) {
                return c;
            }
            int idx = 0;
            int alpha = 255;
            if (data.length == 4) {
                alpha = data[idx++] & 0xFF;
            }
            int r = data[idx++] & 0xFF;
            int g = data[idx++] & 0xFF;
            int b = data[idx++] & 0xFF;

            java.awt.Color color =new java.awt.Color(r, g, b, alpha);
            c.setRGB(data);
            return c;
        } else {
            throw new IllegalStateException();
        }
    }


    public static XSSFColor getFillBackgroundColor(XSSFCellStyle xcs) {
        return getColor(xcs.getFillBackgroundColorColor());

    }

    public static XSSFColor getFillForegroundColor(XSSFCellStyle xcs) {
        return getColor(xcs.getFillForegroundColorColor());

    }

    public static XSSFColor getLeftBorderColor(XSSFCellStyle xcs) {
        return getColor(xcs.getBorderColor(XSSFCellBorder.BorderSide.LEFT));

    }

    public static XSSFColor getRightBorderColor(XSSFCellStyle xcs) {
        return getColor(xcs.getBorderColor(XSSFCellBorder.BorderSide.RIGHT));

    }

    public static XSSFColor getTopBorderColor(XSSFCellStyle xcs) {
        return getColor(xcs.getBorderColor(XSSFCellBorder.BorderSide.TOP));

    }

    public static XSSFColor getBottomBorderColor(XSSFCellStyle xcs) {
        return getColor(xcs.getBorderColor(XSSFCellBorder.BorderSide.BOTTOM));

    }

    private static byte[] getRgbWithTint(XSSFColor c) {
        byte[] rgb = c.getCTColor().getRgb();
        double tint = c.getTint();
        if (rgb != null && tint != 0.0) {
            if(rgb.length == 4) {
                byte[] tmp = new byte[3];
                System.arraycopy(rgb, 1, tmp, 0, 3);
                rgb = tmp;
            }
            for (int i=0; i<rgb.length; i++) {
                int lum = rgb[i] & 0xFF;
                double d = sRGB_to_scRGB(lum / 255.0);
                d = tint > 0 ? d * (1.0 - tint) + tint : d * (1 + tint);
                d = scRGB_to_sRGB(d);
                rgb[i] = (byte)Math.round(d * 255.0);
            }
        }
        return rgb;
    }

    private static double sRGB_to_scRGB(double value) {
        if (value < 0.0) {
            return 0.0;
        }
        if (value <= 0.04045) {
            return value /12.92;
        }
        if (value <= 1.0) {
            return Math.pow(((value + 0.055) / 1.055), 2.4);
        }
        return 1.0;
    }

    private static double scRGB_to_sRGB(double value) {
        if (value < 0.0) {
            return 0.0;
        }
        if (value <= 0.0031308) {
            return value * 12.92;
        }
        if (value < 1.0) {
            return 1.055 * (Math.pow(value, (1.0 / 2.4))) - 0.055;
        }
        return 1.0;
    }



    private static int getColorRGB(XSSFColor c) {
        if (c == null) {
            return 0;
        }else if (c instanceof XSSFColor) {
            XSSFColor xc = c;
            byte[] data = null;
            if (xc.getTint() != 0.0) {
                data = getRgbWithTint(xc);
            } else {
                data = xc.getRGB();
            }
            if (data == null) {
                return 0;
            }
            int idx = 0;
            int alpha = 255;
            if (data.length == 4) {
                alpha = data[idx++] & 0xFF;
            }
            int r = data[idx++] & 0xFF;
            int g = data[idx++] & 0xFF;
            int b = data[idx++] & 0xFF;

            java.awt.Color color =new java.awt.Color(r, g, b, alpha);
            return color.getRGB();
        } else {
            throw new IllegalStateException();
        }
    }


    public static int getFillBackgroundColorRGB(XSSFCellStyle xcs) {
        return getColorRGB(xcs.getFillBackgroundColorColor());

    }

    public static int getFillForegroundColorRGB(XSSFCellStyle xcs) {
        return getColorRGB(xcs.getFillForegroundColorColor());

    }

    public static int getLeftBorderColorRGB(XSSFCellStyle xcs) {
        return getColorRGB(xcs.getBorderColor(XSSFCellBorder.BorderSide.LEFT));

    }

    public static int getRightBorderColorRGB(XSSFCellStyle xcs) {
        return getColorRGB(xcs.getBorderColor(XSSFCellBorder.BorderSide.RIGHT));

    }

    public static int getTopBorderColorRGB(XSSFCellStyle xcs) {
        return getColorRGB(xcs.getBorderColor(XSSFCellBorder.BorderSide.TOP));

    }

    public static int getBottomBorderColorRGB(XSSFCellStyle xcs) {
        return getColorRGB(xcs.getBorderColor(XSSFCellBorder.BorderSide.BOTTOM));

    }


    public Map<String,Object> excelReadAndMappingSheetBean(FileItem fileitem , Map<String,Object> param){


        String emailWorkTargId = "";
        String emailSndLogId = "";
        int getExcelSheetCnt = 0 ;
        Workbook sourceWorkBook = null;


        //-- Map<String,Map<String,Map<String,Object>>> sheetDataMap --
        //sheetName -> key
        //rowDataMap -> value
        Map<String,Map<String,Map<String,Object>>> sheetDataMap = new HashMap<String, Map<String, Map<String, Object>>>();

        Map<String,Object> dataMap = new HashMap<String, Object>();

        try{

            //메일에서 받아온 파일을 읽어들인다.
            if(fileitem.getFile().exists()){


                // 화면단에서 정의한 attachment의 file grp_cd를 가지고 취득하여 excel을 가져온다.
                sourceWorkBook = new XSSFWorkbook(OPCPackage.open(fileitem.getFile().getPath()));

                SheetInfoBean sheetInfo = new SheetInfoBean();

                //해당 엑셀 파일 내에 존재하는 Sheet의 갯수를 가져온다.
                getExcelSheetCnt = sourceWorkBook.getNumberOfSheets();

                // Excel에 대한 데이터를 가져온다.
                XSSFSheet source = ((XSSFWorkbook) sourceWorkBook).getSheetAt(0); //첫번째 Sheet / row를 기준으로 생각한다.
                XSSFRow srcRow = source.getRow(0);

                XSSFCell emailWorkTargIdNewCell = srcRow.getCell(1); // n
                XSSFCell emailSndLogIdNewCell = srcRow.getCell(2);



                // EMAIL_WORK_TARG_ID
                if( null != emailWorkTargIdNewCell){
                    emailWorkTargId = emailWorkTargIdNewCell.getStringCellValue() == null? "" : emailWorkTargIdNewCell.getStringCellValue();
                }

                // EMAIL_SND_LOG_ID
                if( null != emailSndLogIdNewCell){
                    emailSndLogId = emailSndLogIdNewCell.getStringCellValue() == null? "" : emailSndLogIdNewCell.getStringCellValue();
                }

                //주요키값을 가져온다. 해당 키값을 가지고, EASMWSH에서 TEMP_FILE을 조회해야한다.. ( mail send 이전 template file , data와 매칭 시켜야함. )
                if(StringUtils.isEmpty(emailWorkTargId)){
                    try{
                        throw new FileNotFoundException("메일 EXCEL 내에 EMAIL_WORK_TARG_ID 가 존재하지 않습니다.");
                    }catch (Exception e){
                        LOG.error(e.getMessage());
                    }
                }

                if(StringUtils.isEmpty(emailSndLogId)){
                    try{
                        throw new FileNotFoundException("메일 EXCEL 내에 EMAIL_SND_LOG_ID 가 존재하지 않습니다.");
                    }catch (Exception e){
                        LOG.error(e.getMessage());
                    }
                }
            }else{
                try{
                    throw new FileNotFoundException("메일 EXCEL File이 서버에 존재하지 않습니다.");
                }catch (Exception e){
                    LOG.error(e.getMessage());
                }
            }


            String getParamEmailSndLogId = param.get("email_snd_log_id") == null? "" : param.get("email_snd_log_id").toString();
            if(!emailSndLogId.equals(getParamEmailSndLogId)){
                try{
                    throw new FileNotFoundException("메일 EXCEL 내에 EMAIL_SND_LOG_ID와 메일의 EMAIL_SND_LOG_ID가 다릅니다. ");
                }catch (Exception e){
                    LOG.error(e.getMessage());
                }
            }

            Map<String,Object> resultGetExcelFindKey = new HashMap<String, Object>();
            resultGetExcelFindKey.put("email_work_targ_id",emailWorkTargId);
            resultGetExcelFindKey.put("email_snd_log_id",emailSndLogId);

            //가져온 키 값으로, 조회한다. (temp_file_template 확인용 )
            Map<String,Object> sendMailInfo = sqlSession.selectOne("mailWork.mailReceiveExcelFileGetInfo",resultGetExcelFindKey);


            //미리 만들어둔 temp file grp_cd
            String attNo = sendMailInfo.get("tmp_form_file") == null? "" : sendMailInfo.get("tmp_form_file").toString();


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
                try{
                    throw new FileNotFoundException("엑셀파일이 없습니다");
                }catch (Exception e){
                    LOG.error(e.getMessage());
                }
            }

            try {
                excelFileItem = fileService.findDownloadItem(excelFileItem.getId());
            } catch (Exception e1) {
                try{
                    throw new FileNotFoundException("파일을 가져오는 중 오류발생!");
                }catch (Exception e){
                    LOG.error(e.getMessage());
                }

            }


            // 메일로 받은 Excel에 대해서 SheetList 형태로 변경한다.
            // SheetInfoBean List new
            List<SheetInfoBean> getExcelSheetList = new ArrayList<SheetInfoBean>();

            String sendEmailWorkId = sendMailInfo.get("email_work_id") == null? "" :sendMailInfo.get("email_work_id").toString();
            String sendEmailWorkCd = sendMailInfo.get("email_work_cd") == null? "" :sendMailInfo.get("email_work_cd").toString();

            XSSFWorkbook getExcelWorkBook = new XSSFWorkbook(OPCPackage.open(fileitem.getFile().getPath()));

            // SheetCnt 갯수만큼 시트별 데이터를 bean에 담는다.
            for(int i = 0; i < getExcelSheetCnt; i++) {


                SheetInfoBean sheetInfo = new SheetInfoBean();
                // 화면단에서 정의한 attachment의 file grp_cd를 가지고 취득하여 excel을 가져온다.



                //Sheet에 대한 UUID
                String xlsWorkSht = UUID.randomUUID().toString();

                // Excel에 대한 데이터를 가져온다.
                XSSFSheet source = ((XSSFWorkbook) getExcelWorkBook).getSheetAt(i);

                sheetInfo.setEmail_work_id(sendEmailWorkId);
                sheetInfo.setXls_work_sht(xlsWorkSht);
                sheetInfo.setXls_work_sht_nm(source.getSheetName());



                //Excel에 ROW / CELL 정보를 취합한다.
                List<RowInfoBean> sheetRow = readExcel(source,sendEmailWorkId,xlsWorkSht);

                //EXCEL에서 읽어온 정보를 기준으로 ROW LIST를 SET 한다.
                sheetInfo.setRowList(sheetRow);

                //ArrayList add
                getExcelSheetList.add(sheetInfo);
            }



            // SheetInfoBean List new
            List<SheetInfoBean> templateSheetList = new ArrayList<SheetInfoBean>();

            // SEND 처리간 발송된 Template 구현 내역과 , Cell 위치값을 기준으로 Mail로 받은 파일의 key를 매칭시킨다..
            if(excelFileItem.getFile().exists()){

                //sendmailinfo에 temp file을 담는다.
                sendMailInfo.put("tmpExcelFileItem",excelFileItem);

                // 화면단에서 정의한 attachment의 file grp_cd를 가지고 취득하여 excel을 가져온다.
                Workbook sendMailTemplateWorkBook = new XSSFWorkbook(OPCPackage.open(excelFileItem.getFile().getPath()));

                //해당 엑셀 파일 내에 존재하는 Sheet의 갯수를 가져온다.
                int sheetCnt = sendMailTemplateWorkBook.getNumberOfSheets();

                // SheetCnt 갯수만큼 시트별 데이터를 bean에 담는다.
                for(int i = 0; i < sheetCnt; i++) {

                    SheetInfoBean sheetInfo = new SheetInfoBean();

                    // Excel에 대한 데이터를 가져온다.
                    XSSFSheet source = ((XSSFWorkbook) sendMailTemplateWorkBook).getSheetAt(i);

                    for(SheetInfoBean getSheetInfo : getExcelSheetList){

                        String sourceSheetName = source.getSheetName() == null? "" : source.getSheetName();
                        String getSheetInfoSheetName =  getSheetInfo.getXls_work_sht_nm() == null? "" :  getSheetInfo.getXls_work_sht_nm();

                        //SheetName은 독립적이기에 비교하여, 처리 가능. SheetInfo를 여기서 가져옴.
                        if(sourceSheetName.equals(getSheetInfoSheetName)) {


                            String emailWorkId = getSheetInfo.getEmail_work_id();
                            String xlsWorkSht = getSheetInfo.getXls_work_sht();

                            sheetInfo.setEmail_work_id(sendEmailWorkId);
                            sheetInfo.setXls_work_sht(xlsWorkSht);
                            sheetInfo.setXls_work_sht_nm(source.getSheetName());

                            //Excel에 ROW / CELL 정보를 취합한다.
                            List<RowInfoBean> sheetRow = readExcel(source,sendEmailWorkId,xlsWorkSht);

                            //EXCEL에서 읽어온 정보를 기준으로 ROW LIST를 SET 한다.
                            sheetInfo.setRowList(sheetRow);

                            break;
                        }
                    }



                    //Excel에 ROW / CELL 정보를 취합한다.
                    List<RowInfoBean> sheetRow = readExcel(source,sheetInfo.getEmail_work_id(),sheetInfo.getXls_work_sht());

                    //EXCEL에서 읽어온 정보를 기준으로 ROW LIST를 SET 한다.
                    sheetInfo.setRowList(sheetRow);

                    //ArrayList add
                    templateSheetList.add(sheetInfo);
                }
            }

            Map<String,Map<String,Object>> subMap = new HashMap<String, Map<String, Object>>(); //subMap으로 ${dataBean.key} 로 구성될 경우 찾아오기위한 map


            //getExcelSheetList & templateSheetList 비교 시작
            for(SheetInfoBean tempSheetInfo : templateSheetList){

                //sheetInfo를 기준으로 동일 sheet인지 구분.
                SheetInfoBean getSheetInfo = new SheetInfoBean();

                String tempSheetName = tempSheetInfo.getXls_work_sht_nm() == null? "" : tempSheetInfo.getXls_work_sht_nm();

                for(SheetInfoBean sheetInfo : getExcelSheetList){
                    String getSheetInfoSheetName =  sheetInfo.getXls_work_sht_nm() == null? "" :  sheetInfo.getXls_work_sht_nm();

                    //SheetName은 독립적이기에 비교하여, 처리 가능. SheetInfo를 여기서 가져옴.
                    if(tempSheetName.equals(getSheetInfoSheetName)) {
                        BeanUtils.copyProperties(sheetInfo,getSheetInfo);
                        break;
                    }
                }


                //tempSheetInfo와 getSheetInfo를 기준으로 비교하여, value와 key를 매칭 시켜야함.
                List<RowInfoBean> templateRowList = tempSheetInfo.getRowList();
                List<RowInfoBean> getRowList = getSheetInfo.getRowList();

                //-- Map<String,Map<String,Object>> rowDataMap --
                //rowNo ->key
                //dataMap -> value
                Map<String,Map<String,Object>> rowDataMap = new HashMap<String, Map<String, Object>>();

                int listRowNoCheck = 0;

                Map<String,List<Map<String,Object>>> dataListMap = new HashMap<String, List<Map<String, Object>>>(); // $[ ] list data를 map형태로 객체화 시키기 위한 map

                //row별로 찾는다.
                for(RowInfoBean templateRowInfo : templateRowList){

                    int tempRowNo = templateRowInfo.getRow_no();
                    RowInfoBean getRowInfo = new RowInfoBean();

                    for(RowInfoBean getRow : getRowList){

                        int getRowNo = getRow.getRow_no();

                        //rowno 기준으로 동일 시 판단 한다.
                        if(tempRowNo == getRowNo){
                            BeanUtils.copyProperties(getRow,getRowInfo);
                            break;
                        }
                    }

                    if(listRowNoCheck == 0) listRowNoCheck = tempRowNo; // map형태로 취득하던 내역을 list 일 경우, row no 을 체크해서 map -> listMap으로 변환한다.

                    List<CellInfoBean> templateCellList = templateRowInfo.getCellList();
                    List<CellInfoBean> getCellList = getRowInfo.getCellList();

                    //-- Map<String,Object> dataMap --
                    //templateCell string value -> key
                    //getcell string value -> value
                   // Map<String,Object> dataMap = new HashMap<String, Object>();




                    for(CellInfoBean templateCellInfo : templateCellList){

                        int cellIndex = templateCellInfo.getCell_index();

                        CellInfoBean getCellInfo = new CellInfoBean();

                        for(CellInfoBean getCell : getCellList){

                            int getCellIndex = getCell.getCell_index();

                            //rowno 기준으로 동일 시 판단 한다.
                            if(cellIndex == getCellIndex){
                                BeanUtils.copyProperties(getCell,getCellInfo);
                                break;
                            }
                        }


                        String templateStringValue = templateCellInfo.getString_value();

                        String getStringValue = "";

                        switch (getCellInfo.getCell_type()) {
                            case 1:
                                getStringValue = getCellInfo.getString_value();
                                break;
                            case 0:
                                Integer getDoubleValue = getCellInfo.getDouble_value() == 0? 0: (int)getCellInfo.getDouble_value();
                                getStringValue = Integer.toString(getDoubleValue);
                                break;
                            default:
                                break;
                        }


                        //CellDataMap에 담기.
                        if(!EdocStringUtil.isEmpty(templateStringValue)) {


                            if (templateStringValue.startsWith(ExcelReaderUtil.STARTFORMULATOKEN) && templateStringValue.endsWith(ExcelReaderUtil.ENDFORMULATOKEN)) { //list value인지 확인.
                                templateStringValue = templateStringValue.replace("$[","");
                                templateStringValue = templateStringValue.replace("]","");


                                if(templateStringValue.indexOf(".") > -1){ //  data.value  형식의 구조라면, 아래와 같이.

                                    String dataGetKey = "";
                                    String subDataGetKey = "";
                                    String[] splitValue = templateStringValue.split("\\.");
                                    dataGetKey = splitValue[0];
                                    subDataGetKey = splitValue[1];

                                    Map<String,Object> dumyMap = new HashMap<String, Object>();
                                    if(subMap.size() > 0){ // subMap에 현재 데이터가 들어갓을 경우, (list인 경우 row no 체크해주어야 함 - 2019.08.16)
                                        dumyMap = (subMap.get(dataGetKey) == null || listRowNoCheck != tempRowNo)? new HashMap<String, Object>() : subMap.get(dataGetKey);
                                        dumyMap.put(subDataGetKey,getStringValue);
                                    }else{
                                        dumyMap.put(subDataGetKey,getStringValue);
                                    }

                                    subMap.put(dataGetKey,dumyMap); //sub map


                                    List<Map<String,Object>> dumyListMap = new ArrayList<Map<String, Object>>();
                                    if(listRowNoCheck != tempRowNo){
                                        if(dataListMap.size() > 0){
                                            dumyListMap = dataListMap.get(dataGetKey) == null ? new ArrayList<Map<String, Object>>() : dataListMap.get(dataGetKey);
                                            dumyListMap.add(dumyMap);
                                            dataListMap.put(dataGetKey,dumyListMap);
                                        }else{
                                            dumyListMap.add(dumyMap);
                                            dataListMap.put(dataGetKey,dumyListMap);
                                        }

                                        listRowNoCheck = tempRowNo;
                                    }

                                }else {
                                    dataMap.put(templateStringValue,getStringValue);
                                }
                            }else{

                                if (templateStringValue.startsWith(ExcelReaderUtil.STARTEXPRESSIONTOKEN) && templateStringValue.endsWith(ExcelReaderUtil.ENDEXPRESSIONTOKEN)) {
                                    templateStringValue = templateStringValue.replace("${","");
                                    templateStringValue = templateStringValue.replace("}","");
                                    templateStringValue = templateStringValue.replace("$[","");
                                    templateStringValue = templateStringValue.replace("]","");

                                    if(templateStringValue.indexOf(".") > -1){ //  data.value  형식의 구조라면, 아래와 같이.

                                        String dataGetKey = "";
                                        String subDataGetKey = "";
                                        String[] splitValue = templateStringValue.split("\\.");
                                        dataGetKey = splitValue[0];
                                        subDataGetKey = splitValue[1];

                                        Map<String,Object> dumyMap = new HashMap<String, Object>();
                                        if(subMap.size() > 0){ // subMap에 현재 데이터가 들어갓을 경우,
                                            dumyMap = subMap.get(dataGetKey) == null ? new HashMap<String, Object>() : subMap.get(dataGetKey);
                                            dumyMap.put(subDataGetKey,getStringValue);
                                        }else{
                                            dumyMap.put(subDataGetKey,getStringValue);
                                        }

                                        subMap.put(dataGetKey,dumyMap); //sub map

                                        dataMap.put(dataGetKey,dumyMap);
                                    }
                                }else {
                                    templateStringValue = templateStringValue.replace("${","");
                                    templateStringValue = templateStringValue.replace("}","");
                                    templateStringValue = templateStringValue.replace("$[","");
                                    templateStringValue = templateStringValue.replace("]","");
                                    dataMap.put(templateStringValue,getStringValue);
                                }
                            }


                            if(dataListMap.size() > 0) {
                                dataMap.put("list",dataListMap);
                            }



                        }
                    }

                    //CellDataMap에 다 담기면, RowNo을 Key값으로 다시 담는다.
                    //rowDataMap.put(String.valueOf(tempRowNo),dataMap);
                }

                //RowDataMap을 Sheet별로 담는다.
                //sheetDataMap.put(tempSheetName,rowDataMap);
            }


        dataMap.put("email_snd_log_id",emailSndLogId); //발송메일 UUID
        dataMap.put("email_work_cd",sendEmailWorkCd); //메일 업무코드
        dataMap.put("sendEmailInfo",sendMailInfo);  //발송메일 Info 정보
        dataMap.put("receivedEmailInfo",param);   //수신메일 정보


        }catch (RuntimeException rune){
            LOG.error(rune.getMessage());
        }catch (Exception e){
            LOG.error(e.getMessage());
        }

        //return sheetDataMap;
        return dataMap;

    }


    public static Connection getConnection()
    {
        Connection conn = null;
        try {
            String user = "srm9dv";
            String pw = "srm9dv";
            String url = "jdbc:oracle:thin:@175.124.141.220:1521:emro";

            Class.forName("oracle.jdbc.driver.OracleDriver");
            conn = DriverManager.getConnection(url, user, pw);

            System.out.println("Database에 연결되었습니다.\n");

        } catch (ClassNotFoundException cnfe) {
            System.out.println("DB 드라이버 로딩 실패 :"+cnfe.toString());
        } catch (SQLException sqle) {
            System.out.println("DB 접속실패 : "+sqle.toString());
        } catch (Exception e) {
            System.out.println("Unkonwn error");
            e.printStackTrace();
        }
        return conn;
    }




    public static void main(String[] args) throws IOException, InvalidFormatException {


        fileReadAndTranse();


        String isSurveyDir = "C:\\project\\survey";
        String isTobeDir = isSurveyDir+File.separator+"to-be";

        try{

            for (File info : FileUtils.listFilesAndDirs(new File(isTobeDir), TrueFileFilter.INSTANCE, TrueFileFilter.INSTANCE)) {
                if(info.isDirectory()) {

                    ArrayList<File> getFiles = new ArrayList<File>();

                    arrayDirInFileList(getFiles, info);

                    for (File inDirIsFile : getFiles) {
                        String toBeFileName = inDirIsFile.getName();
                        excelReadAndDBInsert(isTobeDir+File.separator+toBeFileName);
                    }
                }
            }



            System.out.println("check");

        }catch (Exception e){
            System.out.println(e.getMessage());
        }



    }


    public static void excelReadAndDBInsert(String filePath)  throws IOException, InvalidFormatException{

        try{

            Workbook sourceWorkBook = new XSSFWorkbook(OPCPackage.open(filePath));
            // Workbook sourceWorkBook = new XSSFWorkbook(OPCPackage.open("C:\\solution\\tech_surv.xlx"));
            //Workbook sourceWorkBook = new XSSFWorkbook(OPCPackage.open("C:\\solution\\testfile\\test_2.xlsx"));
            //Workbook destinationWorkBook = new XSSFWorkbook(OPCPackage.open("C:\\solution\\testfile\\testFile_temp.xlsx"));
      /*  wb_1.setSheetName(0, "Actual");
        wb_1.createSheet("Last Week");*/

            /*Get sheets from the temp file*/
            XSSFSheet source = ((XSSFWorkbook) sourceWorkBook).getSheetAt(0);
            //XSSFSheet destination = ((XSSFWorkbook) destinationWorkBook).getSheetAt(0);

            List<RowInfoBean> list = readExcel(source,UUID.randomUUID().toString(),UUID.randomUUID().toString());


            /**
             * row 2 ~ 5  사용자 내역
             *
             * USER TABLE
             *
             * USER_ID(PK)
             * NAME
             * DEPT
             * POS
             * HIREDATE
             * TOTAL_CAREER
             * DEV_CAREER
             * MANAGER_CAREER
             * DOMAIN_CAREER
             *
             *
             *  (USER_ID -> UUID)
             *  ROW 2 / CELL LIST 2 => 부서
             *  ROW 2 / CELL LIST 5 => 총경력
             *  ROW 3 / CELL LIST 2 => 이름
             *  ROW 3 / CELL LIST 5 => 개발경력
             *  ROW 4 / CELL LIST 2 => 직급
             *  ROW 4 / CELL LIST 5 => 관리경력
             *  ROW 5 / CELL LIST 2 => 입사년월
             *  ROW 5 / CELL LIST 5 => 도메인경력
             */

            UserInfoBean userInfoBean = new UserInfoBean();

            String type = StringUtils.isEmpty(list.get(0).getCellList().get(1).getString_value()) ? "xlsx" : "xlx";

            if(("xlx").equals(type)){
                userInfoBean.setDEPT(list.get(0).getCellList().get(1).getString_value());
                userInfoBean.setTOTAL_CAREER(list.get(0).getCellList().get(4).getString_value());

                userInfoBean.setNAME(list.get(1).getCellList().get(1).getString_value());
                userInfoBean.setDEV_CAREER(list.get(1).getCellList().get(4).getString_value());

                userInfoBean.setPOS(list.get(2).getCellList().get(1).getString_value());
                userInfoBean.setMANAGER_CAREER(list.get(2).getCellList().get(4).getString_value());

                userInfoBean.setHIREDATE(list.get(3).getCellList().get(1).getString_value());
                userInfoBean.setDOMAIN_CAREER(list.get(3).getCellList().get(4).getString_value());
            }else{
                userInfoBean.setDEPT(list.get(1).getCellList().get(1).getString_value());
                userInfoBean.setTOTAL_CAREER(list.get(1).getCellList().get(4).getString_value());

                userInfoBean.setNAME(list.get(2).getCellList().get(1).getString_value());
                userInfoBean.setDEV_CAREER(list.get(2).getCellList().get(4).getString_value());

                userInfoBean.setPOS(list.get(3).getCellList().get(1).getString_value());
                userInfoBean.setMANAGER_CAREER(list.get(3).getCellList().get(4).getString_value());

                userInfoBean.setHIREDATE(list.get(4).getCellList().get(1).getString_value());
                userInfoBean.setDOMAIN_CAREER(list.get(4).getCellList().get(4).getString_value());
            }



            System.out.println("User Info Set--"+userInfoBean.getNAME());
            System.out.println("User Info file path--"+filePath);



            /**
             * ROW 9 ~ 156 기술 내역
             *
             * CELL 1 => 대분류
             * CELL 2 => 중분류
             * CELL 3 => 단위기술
             * CELL 4 => 직접입력
             * CELL 5 => O DESCRIPTION
             * CELL 6 => 1 DESCRIPTION
             * CELL 7 => 2 DESCRIPTION
             * CELL 8 => 3 DESCRIPTION
             * CELL 9 => 4 DESCRIPTION
             * CELL 10 => 5 DESCRIPTION
             * CELL 11 => 6 DESCRIPTION
             * CELL 12 => 7 DESCRIPTION
             * CELL 13 => 8 DESCRIPTION
             * CELL 14 => 9 DESCRIPTION
             *
             */

            //DB INSERT

            List<TechDescriptionInfoBean> techDescriptionList = new ArrayList<TechDescriptionInfoBean>();


            if(("xlx").equals(type)){
                for(int i=6; i < 153; i++){
                    TechDescriptionInfoBean techDescriptionInfoBean = new TechDescriptionInfoBean();


                    //기술 카테고리

                    DescriptionBean descriptionBean = new DescriptionBean();
                    int asTechInfoCnt = techDescriptionList.size();
                    if(techDescriptionList.size() > 0){
                        asTechInfoCnt = techDescriptionList.size() -1;
                    }


                    String category = StringUtils.isEmpty(list.get(i).getCellList().get(0).getString_value())? techDescriptionList.get(asTechInfoCnt).getDescriptionBean().getCATEGORY() : list.get(i).getCellList().get(0).getString_value();

                    descriptionBean.setCATEGORY(category);


                    String division = "";
                    /**
                     * 모델링	    ""
                     * ""           ""
                     * 프레임워크	Server
                     *  ""          UI
                     * OS           ""
                     *
                     * category division
                     *
                     *  1 row category 만 있으면 category = division
                     *  2 row category 없으면, 이전 row의 category 가져오고, division도 없으면 category = division
                     *  3 row category 있으면, category 들어가고, division있으면 division 넣고,
                     *  4 row category 없으면 이전 row의 category 넣고, division있으면, division 넣고
                     *  5 row category 있으면, category 들어가고, division 없으면, 이전 row와 category 비교해서 일치하지 않으면, division은 현재 row category = division
                     */

                    if(StringUtils.isEmpty(list.get(i).getCellList().get(1).getString_value())){
                        if(techDescriptionList.size() > 0){
                            if( StringUtils.isEmpty(techDescriptionList.get(asTechInfoCnt).getDescriptionBean().getDIVISION())){
                                division = category;
                            }else if(category.equals(list.get(i).getCellList().get(0).getString_value())){
                                division = category;
                            }else{
                                division = techDescriptionList.get(asTechInfoCnt).getDescriptionBean().getDIVISION();
                            }
                        }else{
                            division = category;
                        }
                    }else{
                       division =  list.get(i).getCellList().get(1).getString_value();
                    }
                    descriptionBean.setDIVISION(division);


                    descriptionBean.setSECTION(list.get(i).getCellList().get(2).getString_value());




                    descriptionBean.setUSER_ID(userInfoBean.getUSER_ID());
                    descriptionBean.setDESCRIPTION(list.get(i).getCellList().get(3).getString_value());
                    descriptionBean.setDESCRIPTION_0(list.get(i).getCellList().get(4).getString_value());
                    descriptionBean.setDESCRIPTION_1(list.get(i).getCellList().get(5).getString_value());
                    descriptionBean.setDESCRIPTION_2(list.get(i).getCellList().get(6).getString_value());
                    descriptionBean.setDESCRIPTION_3(list.get(i).getCellList().get(7).getString_value());
                    descriptionBean.setDESCRIPTION_4(list.get(i).getCellList().get(8).getString_value());
                    descriptionBean.setDESCRIPTION_5(list.get(i).getCellList().get(9).getString_value());
                    descriptionBean.setDESCRIPTION_6(list.get(i).getCellList().get(10).getString_value());
                    descriptionBean.setDESCRIPTION_7(list.get(i).getCellList().get(11).getString_value());
                    descriptionBean.setDESCRIPTION_8(list.get(i).getCellList().get(12).getString_value());

                    if(list.get(i).getCellList().size() > 14){
                        String description9 = list.get(i).getCellList().get(13).getString_value() == null? "" : list.get(i).getCellList().get(13).getString_value();
                        descriptionBean.setDESCRIPTION_9(description9);
                    }


                    techDescriptionInfoBean.setDescriptionBean(descriptionBean);

                    techDescriptionList.add(techDescriptionInfoBean);
                }
            }else{
                for(int i=8; i < 155; i++){
                    TechDescriptionInfoBean techDescriptionInfoBean = new TechDescriptionInfoBean();

                    //기술 카테고리

                    DescriptionBean descriptionBean = new DescriptionBean();
                    int asTechInfoCnt = techDescriptionList.size();
                    if(techDescriptionList.size() > 0){
                        asTechInfoCnt = techDescriptionList.size() -1;
                    }


                    String category = StringUtils.isEmpty(list.get(i).getCellList().get(0).getString_value())? techDescriptionList.get(asTechInfoCnt).getDescriptionBean().getCATEGORY() : list.get(i).getCellList().get(0).getString_value();

                    descriptionBean.setCATEGORY(category);


                    String division = "";
                    /**
                     * 모델링	    ""
                     * ""           ""
                     * 프레임워크	Server
                     *  ""          UI
                     * OS           ""
                     *
                     * category division
                     *
                     *  1 row category 만 있으면 category = division
                     *  2 row category 없으면, 이전 row의 category 가져오고, division도 없으면 category = division
                     *  3 row category 있으면, category 들어가고, division있으면 division 넣고,
                     *  4 row category 없으면 이전 row의 category 넣고, division있으면, division 넣고
                     *  5 row category 있으면, category 들어가고, division 없으면, 이전 row와 category 비교해서 일치하지 않으면, division은 현재 row category = division
                     */

                    if(StringUtils.isEmpty(list.get(i).getCellList().get(1).getString_value())){
                        if(techDescriptionList.size() > 0){
                            if( StringUtils.isEmpty(techDescriptionList.get(asTechInfoCnt).getDescriptionBean().getDIVISION())){
                                division = category;
                            }else if(category.equals(list.get(i).getCellList().get(0).getString_value())){
                                division = category;
                            }else{
                                division = techDescriptionList.get(asTechInfoCnt).getDescriptionBean().getDIVISION();
                            }
                        }else{
                            division = category;
                        }
                    }else{
                        division =  list.get(i).getCellList().get(1).getString_value();
                    }
                    descriptionBean.setDIVISION(division);


                    descriptionBean.setSECTION(list.get(i).getCellList().get(2).getString_value());




                    descriptionBean.setUSER_ID(userInfoBean.getUSER_ID());
                    descriptionBean.setDESCRIPTION(list.get(i).getCellList().get(3).getString_value());
                    descriptionBean.setDESCRIPTION_0(list.get(i).getCellList().get(4).getString_value());
                    descriptionBean.setDESCRIPTION_1(list.get(i).getCellList().get(5).getString_value());
                    descriptionBean.setDESCRIPTION_2(list.get(i).getCellList().get(6).getString_value());
                    descriptionBean.setDESCRIPTION_3(list.get(i).getCellList().get(7).getString_value());
                    descriptionBean.setDESCRIPTION_4(list.get(i).getCellList().get(8).getString_value());
                    descriptionBean.setDESCRIPTION_5(list.get(i).getCellList().get(9).getString_value());
                    descriptionBean.setDESCRIPTION_6(list.get(i).getCellList().get(10).getString_value());
                    descriptionBean.setDESCRIPTION_7(list.get(i).getCellList().get(11).getString_value());
                    descriptionBean.setDESCRIPTION_8(list.get(i).getCellList().get(12).getString_value());

                    if(list.get(i).getCellList().size() > 14){
                        String description9 = list.get(i).getCellList().get(13).getString_value() == null? "" : list.get(i).getCellList().get(13).getString_value();
                        descriptionBean.setDESCRIPTION_9(description9);
                    }

                    techDescriptionInfoBean.setDescriptionBean(descriptionBean);

                    techDescriptionList.add(techDescriptionInfoBean);
                }
            }


            System.out.println("read excel end");




            Connection conn = null;
            java.sql.Statement st = null; //DB와 소통하는 통로
            PreparedStatement pstmt = null;
            ResultSet rs = null; //결과 받아서 처리할때
            String userSql = "";
            try {

                conn = getConnection();
                st = conn.createStatement();

                userSql = "INSERT INTO SURVEYUSER (USER_ID,NAME, DEPT,POS,HIREDATE,TOTAL_CAREER,DEV_CAREER,MANAGER_CAREER,DOMAIN_CAREER) VALUES ( ?,?,?,?,?,?,?,?,?)";
                pstmt = conn.prepareStatement(userSql);

                pstmt.setString(1, userInfoBean.getUSER_ID());
                pstmt.setString(2, userInfoBean.getNAME());
                pstmt.setString(3, userInfoBean.getDEPT());
                pstmt.setString(4, userInfoBean.getPOS());
                pstmt.setString(5, userInfoBean.getHIREDATE());
                pstmt.setString(6, userInfoBean.getTOTAL_CAREER());
                pstmt.setString(7, userInfoBean.getDEV_CAREER());
                pstmt.setString(8, userInfoBean.getMANAGER_CAREER());
                pstmt.setString(9, userInfoBean.getDOMAIN_CAREER());
                pstmt.executeUpdate();


                for(TechDescriptionInfoBean tech : techDescriptionList){


                    userSql = "INSERT INTO SURVEYDESCRIPTION(USER_ID,SURVEY_ID,DESCRIPTION,DESCRIPTION_0,DESCRIPTION_1,DESCRIPTION_2,DESCRIPTION_3,DESCRIPTION_4,DESCRIPTION_5"
                            +",DESCRIPTION_6,DESCRIPTION_7,DESCRIPTION_8,DESCRIPTION_9,CATEGORY,DIVISION,SECTION) VALUES ( ?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";
                    pstmt = conn.prepareStatement(userSql);

                    DescriptionBean descriptionBean = tech.getDescriptionBean();

                    pstmt.setString(1, descriptionBean.getUSER_ID());
                    pstmt.setString(2, descriptionBean.getSURVEY_ID());
                    pstmt.setString(3, descriptionBean.getDESCRIPTION());
                    pstmt.setString(4, descriptionBean.getDESCRIPTION_0());
                    pstmt.setString(5, descriptionBean.getDESCRIPTION_1());
                    pstmt.setString(6, descriptionBean.getDESCRIPTION_2());
                    pstmt.setString(7, descriptionBean.getDESCRIPTION_3());
                    pstmt.setString(8, descriptionBean.getDESCRIPTION_4());
                    pstmt.setString(9, descriptionBean.getDESCRIPTION_5());
                    pstmt.setString(10, descriptionBean.getDESCRIPTION_6());
                    pstmt.setString(11, descriptionBean.getDESCRIPTION_7());
                    pstmt.setString(12, descriptionBean.getDESCRIPTION_8());
                    pstmt.setString(13, descriptionBean.getDESCRIPTION_9());
                    pstmt.setString(14, descriptionBean.getCATEGORY());
                    pstmt.setString(15, descriptionBean.getDIVISION());
                    pstmt.setString(16, descriptionBean.getSECTION());

                    pstmt.executeUpdate();

                }

            }  catch (SQLException e) {
                System.out.println("DB 연결 실패!");
                e.printStackTrace();
            } finally {
                try {
                    if(rs != null) rs.close();
                    if(st != null) st.close();
                    if(conn != null) conn.close();
                } catch (SQLException e) {
                    e.printStackTrace();
                }
            }
        }catch (Exception e){
            e.printStackTrace();
        }


    }


    public static void excelTranse(File file ,String transeFilePath) throws Exception{


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

                    switch(type){

                        case Cell.CELL_TYPE_BLANK:
                            break;

                        case Cell.CELL_TYPE_NUMERIC:
                            //System.out.print(cell.getNumericCellValue());
                            newCell.setCellValue(cell.getNumericCellValue());
                            break;

                        case Cell.CELL_TYPE_STRING:
                            //System.out.print(cell.getStringCellValue() + "");
                            newCell.setCellValue(cell.getStringCellValue());
                            break;

                        case Cell.CELL_TYPE_ERROR:
                            newCell.setCellErrorValue(cell.getErrorCellValue());
                            break;

                        case Cell.CELL_TYPE_BOOLEAN:
                            newCell.setCellValue( cell.getBooleanCellValue());
                            break;

                        case Cell.CELL_TYPE_FORMULA:
                            //System.out.print(cell.getCellFormula());
                            newCell.setCellFormula(cell.getCellFormula());
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

    public static ArrayList<File> arrayDirInFileList(ArrayList<File> files , File dir){

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


    public static void fileReadAndTranse(){


        String repLine = "";
        int buffer;
        String isSurveyDir = "C:\\project\\survey";

        String isAsisDir = isSurveyDir+File.separator+"as-is";
        String isTobeDir = isSurveyDir+File.separator+"to-be";

        String isXLSXDir = isSurveyDir+File.separator+"xlsx-excel";


        try{

            for (File info : FileUtils.listFilesAndDirs(new File(isAsisDir), TrueFileFilter.INSTANCE, TrueFileFilter.INSTANCE)) {
                if(info.isDirectory()) {

                    ArrayList<File> getFiles = new ArrayList<File>();

                    arrayDirInFileList(getFiles, info);

                    for (File inDirIsFile : getFiles) {
                        System.out.println("FILE NAME ===="+inDirIsFile.getName());

                        String toBeFileName = inDirIsFile.getName();

                        if(toBeFileName.indexOf("xlsx") > -1){
                            toBeFileName = toBeFileName;
                            //excelTranse(inDirIsFile , isXLSXDir+File.separator+toBeFileName);

                        }else if(toBeFileName.indexOf("xls") > -1){
                            toBeFileName = toBeFileName.replaceAll("xls","xlsx");
                        }
                        excelTranse(inDirIsFile , isTobeDir+File.separator+toBeFileName);



                    }
                }
            }



            System.out.println("check");

        }catch (Exception e){

            e.printStackTrace();
        }
    }


    /**
     * USER TABLE
     */
    public static class UserInfoBean{
        public String USER_ID = UUID.randomUUID().toString();
        public String NAME = "";
        public String DEPT = "";
        public String POS = "";
        public String HIREDATE = "";
        public String TOTAL_CAREER = "";
        public String DEV_CAREER = "";
        public String MANAGER_CAREER = "";
        public String DOMAIN_CAREER = "";

        public String getUSER_ID() {
            return USER_ID;
        }

        public void setUSER_ID(String USER_ID) {
            this.USER_ID = USER_ID;
        }

        public String getNAME() {
            return NAME;
        }

        public void setNAME(String NAME) {
            this.NAME = NAME;
        }

        public String getDEPT() {
            return DEPT;
        }

        public void setDEPT(String DEPT) {
            this.DEPT = DEPT;
        }

        public String getPOS() {
            return POS;
        }

        public void setPOS(String POS) {
            this.POS = POS;
        }

        public String getHIREDATE() {
            return HIREDATE;
        }

        public void setHIREDATE(String HIREDATE) {
            this.HIREDATE = HIREDATE;
        }

        public String getTOTAL_CAREER() {
            return TOTAL_CAREER;
        }

        public void setTOTAL_CAREER(String TOTAL_CAREER) {
            this.TOTAL_CAREER = TOTAL_CAREER;
        }

        public String getDEV_CAREER() {
            return DEV_CAREER;
        }

        public void setDEV_CAREER(String DEV_CAREER) {
            this.DEV_CAREER = DEV_CAREER;
        }

        public String getMANAGER_CAREER() {
            return MANAGER_CAREER;
        }

        public void setMANAGER_CAREER(String MANAGER_CAREER) {
            this.MANAGER_CAREER = MANAGER_CAREER;
        }

        public String getDOMAIN_CAREER() {
            return DOMAIN_CAREER;
        }

        public void setDOMAIN_CAREER(String DOMAIN_CAREER) {
            this.DOMAIN_CAREER = DOMAIN_CAREER;
        }
    }



    public static class DescriptionBean{

        public String USER_ID =  "";
        public String SURVEY_ID = UUID.randomUUID().toString();
        public String DESCRIPTION = "";
        public String DESCRIPTION_0 = "";
        public String DESCRIPTION_1 = "";
        public String DESCRIPTION_2 = "";
        public String DESCRIPTION_3 = "";
        public String DESCRIPTION_4 = "";
        public String DESCRIPTION_5 = "";
        public String DESCRIPTION_6 = "";
        public String DESCRIPTION_7 = "";
        public String DESCRIPTION_8 = "";
        public String DESCRIPTION_9 = "";
        public String CATEGORY = "";
        public String DIVISION = "";
        public String SECTION = "";


        public String getCATEGORY() {
            return CATEGORY;
        }

        public void setCATEGORY(String CATEGORY) {
            this.CATEGORY = CATEGORY;
        }

        public String getDIVISION() {
            return DIVISION;
        }

        public void setDIVISION(String DIVISION) {
            this.DIVISION = DIVISION;
        }

        public String getSECTION() {
            return SECTION;
        }

        public void setSECTION(String SECTION) {
            this.SECTION = SECTION;
        }

        public String getUSER_ID() {
            return USER_ID;
        }

        public void setUSER_ID(String USER_ID) {
            this.USER_ID = USER_ID;
        }



        public String getSURVEY_ID() {
            return SURVEY_ID;
        }

        public void setSURVEY_ID(String SURVEY_ID) {
            this.SURVEY_ID = SURVEY_ID;
        }

        public String getDESCRIPTION() {
            return DESCRIPTION;
        }

        public void setDESCRIPTION(String DESCRIPTION) {
            this.DESCRIPTION = DESCRIPTION;
        }

        public String getDESCRIPTION_0() {
            return DESCRIPTION_0;
        }

        public void setDESCRIPTION_0(String DESCRIPTION_0) {
            this.DESCRIPTION_0 = DESCRIPTION_0;
        }

        public String getDESCRIPTION_1() {
            return DESCRIPTION_1;
        }

        public void setDESCRIPTION_1(String DESCRIPTION_1) {
            this.DESCRIPTION_1 = DESCRIPTION_1;
        }

        public String getDESCRIPTION_2() {
            return DESCRIPTION_2;
        }

        public void setDESCRIPTION_2(String DESCRIPTION_2) {
            this.DESCRIPTION_2 = DESCRIPTION_2;
        }

        public String getDESCRIPTION_3() {
            return DESCRIPTION_3;
        }

        public void setDESCRIPTION_3(String DESCRIPTION_3) {
            this.DESCRIPTION_3 = DESCRIPTION_3;
        }

        public String getDESCRIPTION_4() {
            return DESCRIPTION_4;
        }

        public void setDESCRIPTION_4(String DESCRIPTION_4) {
            this.DESCRIPTION_4 = DESCRIPTION_4;
        }

        public String getDESCRIPTION_5() {
            return DESCRIPTION_5;
        }

        public void setDESCRIPTION_5(String DESCRIPTION_5) {
            this.DESCRIPTION_5 = DESCRIPTION_5;
        }

        public String getDESCRIPTION_6() {
            return DESCRIPTION_6;
        }

        public void setDESCRIPTION_6(String DESCRIPTION_6) {
            this.DESCRIPTION_6 = DESCRIPTION_6;
        }

        public String getDESCRIPTION_7() {
            return DESCRIPTION_7;
        }

        public void setDESCRIPTION_7(String DESCRIPTION_7) {
            this.DESCRIPTION_7 = DESCRIPTION_7;
        }

        public String getDESCRIPTION_8() {
            return DESCRIPTION_8;
        }

        public void setDESCRIPTION_8(String DESCRIPTION_8) {
            this.DESCRIPTION_8 = DESCRIPTION_8;
        }

        public String getDESCRIPTION_9() {
            return DESCRIPTION_9;
        }

        public void setDESCRIPTION_9(String DESCRIPTION_9) {
            this.DESCRIPTION_9 = DESCRIPTION_9;
        }
    }


    public static class TechDescriptionInfoBean{
        public DescriptionBean descriptionBean = new DescriptionBean();


        public DescriptionBean getDescriptionBean() {
            return descriptionBean;
        }

        public void setDescriptionBean(DescriptionBean descriptionBean) {
            this.descriptionBean = descriptionBean;
        }
    }


}
