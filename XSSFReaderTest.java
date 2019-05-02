import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import org.springframework.context.annotation.Bean;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;

import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;
import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class XSSFReaderTest {

    @Test
    public void xssfreaderTest() throws Exception{

/*
        try(InputStream in = new FileInputStream(new File("C:\\solution\\testfile\\dynamic_columns_demo.xls"))) {
            try (OutputStream os = new FileOutputStream(new File("C:\\solution\\testfile\\dynamic_columns_demo_out.xls"))) {*/

        List<ExcelSheetCellReaderInfoBean> excelSheetCellReaderInfoBeanList = new ArrayList<ExcelSheetCellReaderInfoBean>();

        ExcelSheetReaderHandlerBean excelSheetReaderHandlerBean = new ExcelSheetReaderHandlerBean();


        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(new File("C:\\solution\\testfile\\testFile.xlsx"));

        // OPCPackge  파일을 i/o 컨테이너를 생성한다고 함.
        OPCPackage opcPackage = xssfWorkbook.getPackage();

        // 메모리 작은 리더기
        XSSFReader xssfReader = new XSSFReader(opcPackage);

        // Sheet 별 Collection 분할해서 가져온다고함.
        XSSFReader.SheetIterator sheetIterator = (XSSFReader.SheetIterator)xssfReader.getSheetsData();

        StylesTable stylesTable = xssfReader.getStylesTable();

        ReadOnlySharedStringsTable stringsTable = new ReadOnlySharedStringsTable(opcPackage);

        List<XSSFCellStyle> xssfCellStylesList = new ArrayList<XSSFCellStyle>();

        short indexSheetNo = 0;

        while (sheetIterator.hasNext()){

            //엑셀의 스타일 정보 담기.
            xssfCellStylesList.add(xssfWorkbook.getCellStyleAt(indexSheetNo));

            List<List<ExcelCellInfoBean>> dataStringList = new ArrayList<List<ExcelCellInfoBean>>();

            ExcelSheetCellReaderInfoBean excelSheetCellReaderInfoBean = new ExcelSheetCellReaderInfoBean();
            InputStream sheetStream = sheetIterator.next();
            InputSource sheetSource = new InputSource(sheetStream);

            SheetListHandlerReader sheetListHandlerReader = new SheetListHandlerReader(dataStringList, 7);

            // Sheet의 행(row) 및 Cell 이벤트를 생성합니다.
            ContentHandler handler = new XSSFCustomSheetXMLHandler(stylesTable, stringsTable, sheetListHandlerReader, false);

            SAXParserFactory saxFactory=SAXParserFactory.newInstance();
            SAXParser saxParser=saxFactory.newSAXParser();
            //sax parser 방식의 xmlReader를 생성
            XMLReader sheetParser = saxParser.getXMLReader();
            //xml reader에 row와 cell 이벤트를 생성하는 핸들러를 설정한 후.
            sheetParser.setContentHandler(handler);

            sheetParser.parse(sheetSource);


            // Excel에서 읽어올때 빼올수있는 정보는 다 넣는다.
            excelSheetCellReaderInfoBean.setColumnCnt(sheetListHandlerReader.columnCnt);
            excelSheetCellReaderInfoBean.setCurrentCol(sheetListHandlerReader.currentCol);
            excelSheetCellReaderInfoBean.setCurrRowNum(sheetListHandlerReader.currRowNum);
            excelSheetCellReaderInfoBean.setHeader(sheetListHandlerReader.header);
            excelSheetCellReaderInfoBean.setRows(sheetListHandlerReader.rows);
            //excelSheetCellReaderInfoBean.setCellInfoBeanList(sheetListHandlerReader.cellInfoBeanList);


            String sheet_name = sheetIterator.getSheetName() == null? "": sheetIterator.getSheetName().toString();
            String sheet_comment = sheetIterator.getSheetComments() == null? "": sheetIterator.getSheetComments().toString();

            excelSheetCellReaderInfoBean.setSheetComment(sheet_comment);
            excelSheetCellReaderInfoBean.setSheetName(sheet_name);

            excelSheetCellReaderInfoBeanList.add(excelSheetCellReaderInfoBean);

            sheetStream.close();

            indexSheetNo++;
        }

        //엑셀 정보 담기
        excelSheetReaderHandlerBean.setExcelSheetCellReaderInfoBeanList(excelSheetCellReaderInfoBeanList);


        excelSheetReaderHandlerBean.setXssfCellStyleList(xssfCellStylesList);

        opcPackage.close();

        SXSSFCreateTest.createXlsx(excelSheetReaderHandlerBean);

    }


    public class SheetListHandlerReader implements XSSFCustomSheetXMLHandler.SheetContentsHandler{
        //header를 제외한 데이터부분
        private List<List<ExcelCellInfoBean>> rows = new ArrayList<List<ExcelCellInfoBean>>();

        //cell 호출시마다 쌓아놓을 1 row List
        private List<ExcelCellInfoBean> row = new ArrayList<ExcelCellInfoBean>();

        //Header 정보를 입력
        private List<ExcelCellInfoBean> header = new ArrayList<ExcelCellInfoBean>();

        //빈 값을 체크하기 위해 사용할 셀번호
        private int currentCol = -1;

        //현재 읽고 있는 Cell의 Col
        private int currRowNum = 0;

        private int columnCnt;



        //외부 Collection 과 배열 Size를 받음
        public SheetListHandlerReader(List<List<ExcelCellInfoBean>> rows, int columnsCnt){
            this.rows = rows;
            this.columnCnt = columnsCnt;
        }


        public void startRow(int rowNum) {
            //empty 값을 체크하기 위한 초기 셋팅값
            this.currentCol = -1;
            this.currRowNum = rowNum;
        }
        public void endRow(int rowNum) {

            ExcelCellInfoBean cellInfoBean = new ExcelCellInfoBean();

          /*  if(rowNum ==0) {
                header = new ArrayList(row);
            } else {
                //헤더의 길이가 현재 로우보다 더 길다면 Cell의 뒷자리가 빈값임으로 해당값만큼 공백
                if(row.size() < header.size()) {
                    for (int i = row.size(); i < header.size(); i++) {
                        ExcelCellInfoBean nullCellInfo = new ExcelCellInfoBean();
                        nullCellInfo.setValue("");
                        row.add(nullCellInfo);
                    }
                }*//*
                rows.add(new ArrayList(row));
            }*/
            rows.add(new ArrayList(row));

            row.clear();
        }

        public void cell(String columnName, String value, XSSFComment var3, XSSFCellStyle cellStyle) {

            int iCol = (new CellReference(columnName)).getCol();

            int emptyCol = iCol - currentCol - 1;

            ExcelCellInfoBean cellInfoBean = new ExcelCellInfoBean();

            //읽은 Cell의 번호를 이용하여 빈Cell 자리에 빈값을 강제로 저장시켜줌
            for(int i = 0 ; i < emptyCol ; i++) {
                ExcelCellInfoBean nullCellInfo = new ExcelCellInfoBean();
                nullCellInfo.setValue("");
                row.add(nullCellInfo);
            }
            currentCol = iCol;
            cellInfoBean.setIndex(currentCol);
            cellInfoBean.setCellStyle(cellStyle);
            cellInfoBean.setValue(value);
            row.add(cellInfoBean);
        }
        public void headerFooter(String text, boolean isHeader, String tagName) {
        }
    }


    /**
     *  임시 CELL의 정보를 담당하는 BEAN
     */
    public class ExcelCellInfoBean{
        public Integer index = null;

        public String value = "";

        public XSSFCellStyle cellStyle = null;


        public Integer getIndex() {
            return index;
        }

        public void setIndex(Integer index) {
            this.index = index;
        }

        public String getValue() {
            return value;
        }

        public void setValue(String value) {
            this.value = value;
        }

        public XSSFCellStyle getCellStyle() {
            return cellStyle;
        }

        public void setCellStyle(XSSFCellStyle cellStyle) {
            this.cellStyle = cellStyle;
        }
    }



    // Excel Reader에서 읽은 내역을 Bean에다가 담아, Sheet별로 정보를 수집한다.
    public class ExcelSheetCellReaderInfoBean{


        //header를 제외한 데이터부분
        public List<List<ExcelCellInfoBean>> rows = new ArrayList<List<ExcelCellInfoBean>>();

        //Header 정보를 입력
        public List<ExcelCellInfoBean> header = new ArrayList<ExcelCellInfoBean>();

        //빈 값을 체크하기 위해 사용할 셀번호
        public int currentCol = -1;

        //현재 읽고 있는 Cell의 Col
        public int currRowNum = 0;

        public Integer columnCnt = null;

        //Excel Sheet Name
        public String sheetName = "";

        //Excel Sheet Comment
        public String sheetComment ="";

        private List<ExcelCellInfoBean> cellInfoBeanList = new ArrayList<ExcelCellInfoBean>();


        public List<List<ExcelCellInfoBean>> getRows() {
            return rows;
        }

        public void setRows(List<List<ExcelCellInfoBean>> rows) {
            this.rows = rows;
        }

        public List<ExcelCellInfoBean> getHeader() {
            return header;
        }

        public void setHeader(List<ExcelCellInfoBean> header) {
            this.header = header;
        }

        public int getCurrentCol() {
            return currentCol;
        }

        public void setCurrentCol(int currentCol) {
            this.currentCol = currentCol;
        }

        public int getCurrRowNum() {
            return currRowNum;
        }

        public void setCurrRowNum(int currRowNum) {
            this.currRowNum = currRowNum;
        }

        public Integer getColumnCnt() {
            return columnCnt;
        }

        public void setColumnCnt(Integer columnCnt) {
            this.columnCnt = columnCnt;
        }

        public String getSheetName() {
            return sheetName;
        }

        public void setSheetName(String sheetName) {
            this.sheetName = sheetName;
        }

        public String getSheetComment() {
            return sheetComment;
        }

        public void setSheetComment(String sheetComment) {
            this.sheetComment = sheetComment;
        }

        public List<ExcelCellInfoBean> getCellInfoBeanList() {
            return cellInfoBeanList;
        }

        public void setCellInfoBeanList(List<ExcelCellInfoBean> cellInfoBeanList) {
            this.cellInfoBeanList = cellInfoBeanList;
        }
    }



    // 엑셀에서 읽은 내역 전체를 EXCEL Generating Class에게 넘겨준다
    public class ExcelSheetReaderHandlerBean{

        public List<ExcelSheetCellReaderInfoBean> excelSheetCellReaderInfoBeanList = new ArrayList<ExcelSheetCellReaderInfoBean>();

        public List<XSSFCellStyle> xssfCellStyleList ;

        public List<ExcelSheetCellReaderInfoBean> getExcelSheetCellReaderInfoBeanList() {
            return excelSheetCellReaderInfoBeanList;
        }

        public void setExcelSheetCellReaderInfoBeanList(List<ExcelSheetCellReaderInfoBean> excelSheetCellReaderInfoBeanList) {
            this.excelSheetCellReaderInfoBeanList = excelSheetCellReaderInfoBeanList;
        }

        public List<XSSFCellStyle> getXssfCellStyleList() {
            return xssfCellStyleList;
        }

        public void setXssfCellStyleList(List<XSSFCellStyle> xssfCellStyleList) {
            this.xssfCellStyleList = xssfCellStyleList;
        }
    }


}


