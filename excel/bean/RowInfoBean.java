package smartsuite.app.common.excel.bean;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

@SuppressWarnings("PMD")
public class RowInfoBean {

    public List<CellInfoBean> cellList = new ArrayList<CellInfoBean>();

    public String row_id = "";

    public String email_work_id = "";

    public String xls_work_sht = "";

    public int row_no = 0;

    public String usrcert_id = "";

    public String auth_id = "";


    public String reg_id = "";

    public Date reg_dt = null;

    public String getAuth_id() {
        return auth_id;
    }

    public void setAuth_id(String auth_id) {
        this.auth_id = auth_id;
    }

    public List<CellInfoBean> getCellList() {
        return cellList;
    }

    public void setCellList(List<CellInfoBean> cellList) {
        this.cellList = cellList;
    }

    public String getRow_id() {
        return row_id;
    }

    public void setRow_id(String row_id) {
        this.row_id = row_id;
    }

    public String getEmail_work_id() {
        return email_work_id;
    }

    public void setEmail_work_id(String email_work_id) {
        this.email_work_id = email_work_id;
    }

    public String getXls_work_sht() {
        return xls_work_sht;
    }

    public void setXls_work_sht(String xls_work_sht) {
        this.xls_work_sht = xls_work_sht;
    }

    public int getRow_no() {
        return row_no;
    }

    public void setRow_no(int row_no) {
        this.row_no = row_no;
    }

    public String getUsrcert_id() {
        return usrcert_id;
    }

    public void setUsrcert_id(String usrcert_id) {
        this.usrcert_id = usrcert_id;
    }

    public String getReg_id() {
        return reg_id;
    }

    public void setReg_id(String reg_id) {
        this.reg_id = reg_id;
    }

    public Date getReg_dt() {
        return reg_dt;
    }

    public void setReg_dt(Date reg_dt) {
        this.reg_dt = reg_dt;
    }
}
