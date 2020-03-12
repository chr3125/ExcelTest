package smartsuite.app.common.excel.bean;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

@SuppressWarnings("PMD")
public class ExcelInfoBean {

    public List<SheetInfoBean> sheetList = new ArrayList<SheetInfoBean>();

    public String email_work_id = "";

    public String tmp_id = "";

    public String email_work_nm = "";

    public String email_work_desc = "";

    public String email_work_cd = "";

    public String use_yn = "Y";

    public String reg_id = "";

    public Date reg_dt = null;

    public String mod_id = "";

    public Date mod_dt = null;

    public String mail_set_id = "";

    public String cnfm_yn = "N";

    public String att_no = "";

    public String getAtt_no() {
        return att_no;
    }

    public void setAtt_no(String att_no) {
        this.att_no = att_no;
    }

    public List<SheetInfoBean> getSheetList() {
        return sheetList;
    }

    public String getCnfm_yn() {
        return cnfm_yn;
    }

    public void setCnfm_yn(String cnfm_yn) {
        this.cnfm_yn = cnfm_yn;
    }

    public String getMail_set_id() {
        return mail_set_id;
    }

    public void setMail_set_id(String mail_set_id) {
        this.mail_set_id = mail_set_id;
    }

    public void setSheetList(List<SheetInfoBean> sheetList) {
        this.sheetList = sheetList;
    }

    public String getEmail_work_id() {
        return email_work_id;
    }

    public void setEmail_work_id(String email_work_id) {
        this.email_work_id = email_work_id;
    }

    public String getTmp_id() {
        return tmp_id;
    }

    public void setTmp_id(String tmp_id) {
        this.tmp_id = tmp_id;
    }

    public String getEmail_work_nm() {
        return email_work_nm;
    }

    public void setEmail_work_nm(String email_work_nm) {
        this.email_work_nm = email_work_nm;
    }

    public String getEmail_work_desc() {
        return email_work_desc;
    }

    public void setEmail_work_desc(String email_work_desc) {
        this.email_work_desc = email_work_desc;
    }

    public String getEmail_work_cd() {
        return email_work_cd;
    }

    public void setEmail_work_cd(String email_work_cd) {
        this.email_work_cd = email_work_cd;
    }

    public String getUse_yn() {
        return use_yn;
    }

    public void setUse_yn(String use_yn) {
        this.use_yn = use_yn;
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

    public String getMod_id() {
        return mod_id;
    }

    public void setMod_id(String mod_id) {
        this.mod_id = mod_id;
    }

    public Date getMod_dt() {
        return mod_dt;
    }

    public void setMod_dt(Date mod_dt) {
        this.mod_dt = mod_dt;
    }
}
