package smartsuite.app.common.excel.bean;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

@SuppressWarnings("PMD")
public class SheetInfoBean {

    public List<RowInfoBean> rowList = new ArrayList<RowInfoBean>();

    public String email_work_id = "";

    public String xls_work_sht = "";

    public String xls_work_sht_nm = "";

    public String reg_id = "";

    public Date reg_dt = null;


    public List<RowInfoBean> getRowList() {
        return rowList;
    }

    public void setRowList(List<RowInfoBean> rowList) {
        this.rowList = rowList;
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

    public String getXls_work_sht_nm() {
        return xls_work_sht_nm;
    }

    public void setXls_work_sht_nm(String xls_work_sht_nm) {
        this.xls_work_sht_nm = xls_work_sht_nm;
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
