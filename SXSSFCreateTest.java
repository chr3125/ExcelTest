import com.fasterxml.jackson.databind.exc.InvalidFormatException;
import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;
import org.junit.Test;
import org.springframework.format.annotation.DateTimeFormat;


import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

public class SXSSFCreateTest {

      /*  public static <T> void createXlsx(List<T> pojoObjectList, String filePath) throws IllegalAccessException, IllegalArgumentException, InvocationTargetException, NoSuchMethodException, SecurityException, IOException {
            Workbook workbook = new SXSSFWorkbook();
            SXSSFSheet sheet = (SXSSFSheet) workbook.createSheet();
            for (int i = 0; i < pojoObjectList.size(); i++) {
                Row row = sheet.createRow(i);
                if (i == 0) {
                    RowFilledWithPojoHeader(pojoObjectList.get(i), row);
                } else {
                    RowFilledWithPojoData(pojoObjectList.get(i), row);
                }
            }
            FileOutputStream fos = new FileOutputStream(filePath + pojoObjectList.get(0).getClass().getSimpleName().toUpperCase() + "_" + System.currentTimeMillis() + ".xlsx");
            workbook.write(fos);
            fos.close();
        }

        private static Row RowFilledWithPojoHeader(Object pojoObject, Row row) throws IllegalAccessException, IllegalArgumentException, InvocationTargetException, NoSuchMethodException, SecurityException {
            Field[] fields = pojoObject.getClass().getDeclaredFields();
            int fieldLength = fields.length;
            for (int i = 0; i < fieldLength; i++) {
                String cellValue = fields[i].getName().toUpperCase();
                row.createCell(i).setCellValue(cellValue);
            }
            return row;
        }

        private static Row RowFilledWithPojoData(Object pojoObject, Row row) throws IllegalAccessException, IllegalArgumentException, InvocationTargetException, NoSuchMethodException, SecurityException {
            Field[] fields = pojoObject.getClass().getDeclaredFields();
            int fieldLength = fields.length;
            for (int i = 0; i < fieldLength; i++) {
                Method method = pojoObject.getClass().getMethod("get" + fields[i].getName().substring(0, 1).toUpperCase() + fields[i].getName().substring(1));
                String cellValue;
                String returnType = method.getReturnType().getName();
                if (returnType.equals("org.joda.time.DateTime")) {
                    Object dateTime = method.invoke(pojoObject);
                    if (dateTime == null) {
                        cellValue = "";
                    } else {
                        SimpleDateFormat sd = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss", Locale.getDefault());
                        cellValue = sd.format(dateTime);
                    }
                } else {
                    cellValue = (String) method.invoke(pojoObject);
                }
                row.createCell(i).setCellValue(cellValue);
            }
            return row;
        }

        public static String getCellValue(Row row, int cellIndex) {
            Cell cell = row.getCell(cellIndex);
            if (cell == null) {
                return null;
            }
            cell.setCellType(Cell.CELL_TYPE_STRING);
            return cell.getStringCellValue();
        }

        public static String getCellValue(Object object) {
            if (object == null) {
                return "";
            }
            return object.toString();
        }
*/

      public static <T> void createXlsx(XSSFReaderTest.ExcelSheetReaderHandlerBean excelsheetreaderhandlerbean) throws Exception {
          InputStream is = new FileInputStream(new File("C:\\solution\\testfile\\testFile_temp.xlsx"));
          OutputStream out = new FileOutputStream(new File("C:\\solution\\testfile\\testFile_out.xlsx"));

          //input stream으로 tempfile을 만든다
          XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);

          // 엑셀템플릿파일에 쓰여질 부분 검색
          /*Sheet originSheet = xssfWorkbook.getSheetAt();
          rowNo = originSheet.getLastRowNum();*/

          //SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook(xssfWorkbook, 300);

          List<XSSFReaderTest.ExcelSheetCellReaderInfoBean> excelSheetCellReaderInfoBeanList = excelsheetreaderhandlerbean.getExcelSheetCellReaderInfoBeanList() == null? null : excelsheetreaderhandlerbean.getExcelSheetCellReaderInfoBeanList();




          for (int i = 0; i < excelSheetCellReaderInfoBeanList.size(); i++) {
              //가져온 데이터의 list 별로 Sheet를 만들기 시작한다.
              XSSFSheet createSheet = xssfWorkbook.getSheetAt(i);
              XSSFCellStyle setWorkBookStyle = xssfWorkbook.createCellStyle();

            /*  Font headerFont = xssfWorkbook.createFont();
              headerFont.setBold(true);
              headerFont.setFontHeightInPoints((short) 14);
              headerFont.setColor(IndexedColors.RED.getIndex());
              headerFont.setFontName("맑은 고딕");

              setWorkBookStyle.setFont(headerFont);*/


              /*XSSFCellStyle orgCellStyle = excelsheetreaderhandlerbean.getXssfCellStyleList().get(i);
              setWorkBookStyle.cloneStyleFrom(orgCellStyle);*/

              // 해당 리스트의 rows의 리스트를 가져온다.
              for(int a = 0; a < excelSheetCellReaderInfoBeanList.get(i).getRows().size(); a++){
                  XSSFRow row = createSheet.createRow(a);
                  List<XSSFReaderTest.ExcelCellInfoBean> excelSheetCellReaderInfoBean = excelSheetCellReaderInfoBeanList.get(i).getRows().get(a);


                  for(int b = 0; b < excelSheetCellReaderInfoBean.size(); b++){
                      XSSFCell cell = row.createCell(b);
                      cell.setCellValue(excelSheetCellReaderInfoBean.get(b).getValue());
                      cell.setCellValue(excelSheetCellReaderInfoBean.get(b).getValue());

                      XSSFCellStyle xssfCellStyle = excelSheetCellReaderInfoBean.get(b).getCellStyle() == null? null : excelSheetCellReaderInfoBean.get(b).getCellStyle();
                      if(null != xssfCellStyle ) {
                          XSSFCellStyle tempCellStyle = cell.getCellStyle();
                          tempCellStyle.cloneStyleFrom(xssfCellStyle);
                          cell.setCellStyle(tempCellStyle);
                      }

                     //cell.setCellStyle(setWorkBookStyle);

                      /*XSSFCellStyle xssfCellStyle = excelSheetCellReaderInfoBean.get(b).getCellStyle() == null? null : excelSheetCellReaderInfoBean.get(b).getCellStyle();
                      if(null != xssfCellStyle ) setWorkBookStyle.cloneStyleFrom(xssfCellStyle);  cell.setCellStyle(xssfCellStyle);*/
                     // cell.setCellStyle(excelSheetCellReaderInfoBean.get(b).getCellStyle());

                  }
              }
              // 디스크로 flush
              //((SXSSFSheet)createSheet).flushRows(excelSheetCellReaderInfoBeanList.get(i).getRows().size());



          }
          //sxssfWorkbook.write(out);
            xssfWorkbook.write(out);
          // 디스크에 임시파일로 저장한 파일 삭제
         // sxssfWorkbook.dispose();
          out.close();
          is.close();
         /* FileOutputStream fos = new FileOutputStream(filePath + pojoObjectList.get(0).getClass().getSimpleName().toUpperCase() + "_" + System.currentTimeMillis() + ".xlsx");
          workbook.write(fos);
          fos.close();*/
      }


}
