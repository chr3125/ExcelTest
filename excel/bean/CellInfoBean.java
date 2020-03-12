package smartsuite.app.common.excel.bean;

import java.util.UUID;

@SuppressWarnings("PMD")
public class CellInfoBean {


    public String cell_id = null;

    public String sheet_id = null;

    public String row_id = null;

    //FONT BOLD
    public boolean bold = false;

    //FONT HEIGHT
    public short font_height = 220;

    //FONT FONT NAME
    public String font_nm = "맑은 고딕";

    //FONT ITALIC
    public boolean italic = false;

    //FONT STRIKEOUT
    public boolean strikeout = false;

    //FONT TYPEOFFSET
    public short type_offset = 0;

    //FONT UNDERLINE
    public byte under_line = 0;

    //FONT CHARSET
    public int charset = 129;

    //FONT COLOR
    public int color = -16777216;

    //FONT DATAFORMAT
    public short dataformat = 0;

    //CELL_STYLE ALIGNMENT_CODE code값을 가지고 다시 찾아와 비지니스 로직에서 박아주는 형태로 진행되어야 할것으로 보임. (정렬 코드)
    /**
     *     int INT_GENERAL = 1;
     *     int INT_LEFT = 2;
     *     int INT_CENTER = 3;
     *     int INT_RIGHT = 4;
     *     int INT_FILL = 5;
     *     int INT_JUSTIFY = 6;
     *     int INT_CENTER_CONTINUOUS = 7;
     *     int INT_DISTRIBUTED = 8;
     */
    public int alignment_cd = 0;

    //CELL_STYLE HIDDEN ( 숨김 여부 )
    public boolean hidden = false;

    //CELL_STYLE LOCKED ( 잠금여부 )
    public boolean locked = false;

    //CELL_STYLE WRAPTEXT (줄바꿈여부)
    public boolean wraptext = false;

    //CELL_STYLE BORDER CODE 4방향 동일 code값을 가지고 다시 찾아와 비지니스 로직에서 박아주는 형태로 진행되어야 할것으로 보임.
    /**
     *     short BORDER_NONE = 0;
     *     short BORDER_THIN = 1;
     *     short BORDER_MEDIUM = 2;
     *     short BORDER_DASHED = 3;
     *     short BORDER_HAIR = 7;
     *     short BORDER_THICK = 5;
     *     short BORDER_DOUBLE = 6;
     *     short BORDER_DOTTED = 4;
     *     short BORDER_MEDIUM_DASHED = 8;
     *     short BORDER_DASH_DOT = 9;
     *     short BORDER_MEDIUM_DASH_DOT = 10;
     *     short BORDER_DASH_DOT_DOT = 11;
     *     short BORDER_MEDIUM_DASH_DOT_DOT = 12;
     *     short BORDER_SLANTED_DASH_DOT = 13;
     */
    public short border_bottom_cd = 1;
    public short border_left_cd = 1;
    public short border_right_cd = 1;
    public short border_top_cd = 1;








    // COLOR 들에 한하여서는 RGB 값을 빼와서.. tiny 값을 기준으로 처리를 할 생각
    //CELL_STYLE BottomBorderColor
    public int bottom_border_color = 0;

    //CELL_STYLE FillBackgroundColor
    public int fill_background_color = 0;

    //CELL_STYLE FillBackgroundColor
    public int fill_foreground_color = 0;

    //CELL_STYLE LeftBorderColor
    public int left_border_color = 0;

    //CELL_STYLE RightBorderColor
    public int right_border_color = 0;

    //CELL_STYLE TopBorderColor
    public int top_border_color = 0;

    //CELL_STYLE Indention
    public short indention = 0;

    //CELL_STYLE Rotation
    public short rotation = 0;

    //CELL_STYLE VerticalAlignment code값을 가지고 다시 찾아와 비지니스 로직에서 박아주는 형태로 진행되어야 할것으로 보임.
    /**
     *     int INT_TOP = 1;
     *     int INT_CENTER = 2;
     *     int INT_BOTTOM = 3;
     *     int INT_JUSTIFY = 4;
     *     int INT_DISTRIBUTED = 5;
     */
    public int vertical_alignment_cd = 2;

    //CELL_TYPE
    /**
     *      STRING: 1
     *      NUMERIC: 0
     *      BLANK: 3
     *      BOOLEAN: 4
     *      ERROR: 5
     *      FORMULA: 2
     *      DEFAULT : 0 -> 해당 코드는 NULL 처리 되어야함.
     */
    public int cell_type = 0;

    //CELL_STYLE FillPattern
    /**
     *     NO_FILL = 0;
     *     SOLID_FOREGROUND = 1;
     *     FINE_DOTS = 2;
     *     ALT_BARS = 3;
     *     SPARSE_DOTS = 4;
     *     THICK_HORZ_BANDS = 5;
     *     THICK_VERT_BANDS = 6;
     *     THICK_BACKWARD_DIAG = 7;
     *     THICK_FORWARD_DIAG = 8;
     *     BIG_SPOTS = 9;
     *     BRICKS = 10;
     *     THIN_HORZ_BANDS = 11;
     *     THIN_VERT_BANDS = 12;
     *     THIN_BACKWARD_DIAG = 13;
     *     THIN_FORWARD_DIAG = 14;
     *     SQUARES = 15;
     *     DIAMONDS = 16;
     */
    public short fill_pattern = 0;


    /**
     *  switch (oldCell.getCellType()) {
     *             case STRING:
     *                 newCell.setCellValue(oldCell.getStringCellValue());
     *                 break;
     *             case NUMERIC:
     *                 newCell.setCellValue(oldCell.getNumericCellValue());
     *                 break;
     *             case BLANK:
     *                 newCell.setCellType(CellType.BLANK);
     *                 break;
     *             case BOOLEAN:
     *                 newCell.setCellValue(oldCell.getBooleanCellValue());
     *                 break;
     *             case ERROR:
     *                 newCell.setCellErrorValue(oldCell.getErrorCellValue());
     *                 break;
     *             case FORMULA:
     *                 newCell.setCellFormula(oldCell.getCellFormula());
     *                 formulaInfoList.add(new FormulaInfo(oldCell.getSheet().getSheetName(), oldCell.getRowIndex(), oldCell
     *                         .getColumnIndex(), oldCell.getCellFormula()));
     *                 break;
     *             default:
     *                 break;
     *         }
     *
     */


    public String string_value ="";

    public double double_value = 0;

    public boolean boolean_value = false;

    public byte error_value = 0;

    public String formula_value = null;

    public int cell_index = 0;

    public int col_width = 0;

    public int getCol_width() {
        return col_width;
    }

    public void setCol_width(int col_width) {
        this.col_width = col_width;
    }

    public short getDataformat() {
        return dataformat;
    }

    public void setDataformat(short dataformat) {
        this.dataformat = dataformat;
    }

    // 개발자가 추가적으로 Cell을 생성하기 위한 플레그
    public boolean appendCellFlag = false;

    public String getCell_id() {
        return cell_id;
    }

    public void setCell_id(String cell_id) {
        this.cell_id = cell_id;
    }

    public String getSheet_id() {
        return sheet_id;
    }

    public void setSheet_id(String sheet_id) {
        this.sheet_id = sheet_id;
    }

    public String getRow_id() {
        return row_id;
    }

    public void setRow_id(String row_id) {
        this.row_id = row_id;
    }

    public boolean isBold() {
        return bold;
    }

    public void setBold(boolean bold) {
        this.bold = bold;
    }

    public short getFont_height() {
        return font_height;
    }

    public void setFont_height(short font_height) {
        this.font_height = font_height;
    }

    public String getFont_nm() {
        return font_nm;
    }

    public void setFont_nm(String font_nm) {
        this.font_nm = font_nm;
    }

    public boolean isItalic() {
        return italic;
    }

    public void setItalic(boolean italic) {
        this.italic = italic;
    }

    public boolean isStrikeout() {
        return strikeout;
    }

    public void setStrikeout(boolean strikeout) {
        this.strikeout = strikeout;
    }

    public short getType_offset() {
        return type_offset;
    }

    public void setType_offset(short type_offset) {
        this.type_offset = type_offset;
    }

    public byte getUnder_line() {
        return under_line;
    }

    public void setUnder_line(byte under_line) {
        this.under_line = under_line;
    }

    public int getCharset() {
        return charset;
    }

    public void setCharset(int charset) {
        this.charset = charset;
    }

    public int getColor() {
        return color;
    }

    public void setColor(int color) {
        this.color = color;
    }



    public boolean isHidden() {
        return hidden;
    }

    public void setHidden(boolean hidden) {
        this.hidden = hidden;
    }

    public boolean isLocked() {
        return locked;
    }

    public void setLocked(boolean locked) {
        this.locked = locked;
    }

    public boolean isWraptext() {
        return wraptext;
    }

    public void setWraptext(boolean wraptext) {
        this.wraptext = wraptext;
    }


    public int getBottom_border_color() {
        return bottom_border_color;
    }

    public void setBottom_border_color(int bottom_border_color) {
        this.bottom_border_color = bottom_border_color;
    }

    public int getFill_background_color() {
        return fill_background_color;
    }

    public void setFill_background_color(int fill_background_color) {
        this.fill_background_color = fill_background_color;
    }

    public int getFill_foreground_color() {
        return fill_foreground_color;
    }

    public void setFill_foreground_color(int fill_foreground_color) {
        this.fill_foreground_color = fill_foreground_color;
    }

    public int getLeft_border_color() {
        return left_border_color;
    }

    public void setLeft_border_color(int left_border_color) {
        this.left_border_color = left_border_color;
    }

    public int getRight_border_color() {
        return right_border_color;
    }

    public void setRight_border_color(int right_border_color) {
        this.right_border_color = right_border_color;
    }

    public int getTop_border_color() {
        return top_border_color;
    }

    public void setTop_border_color(int top_border_color) {
        this.top_border_color = top_border_color;
    }

    public short getIndention() {
        return indention;
    }

    public void setIndention(short indention) {
        this.indention = indention;
    }

    public short getRotation() {
        return rotation;
    }

    public void setRotation(short rotation) {
        this.rotation = rotation;
    }

    public int getAlignment_cd() {
        return alignment_cd;
    }

    public void setAlignment_cd(int alignment_cd) {
        this.alignment_cd = alignment_cd;
    }

    public short getBorder_bottom_cd() {
        return border_bottom_cd;
    }

    public void setBorder_bottom_cd(short border_bottom_cd) {
        this.border_bottom_cd = border_bottom_cd;
    }

    public short getBorder_left_cd() {
        return border_left_cd;
    }

    public void setBorder_left_cd(short border_left_cd) {
        this.border_left_cd = border_left_cd;
    }

    public short getBorder_right_cd() {
        return border_right_cd;
    }

    public void setBorder_right_cd(short border_right_cd) {
        this.border_right_cd = border_right_cd;
    }

    public short getBorder_top_cd() {
        return border_top_cd;
    }

    public void setBorder_top_cd(short border_top_cd) {
        this.border_top_cd = border_top_cd;
    }

    public int getVertical_alignment_cd() {
        return vertical_alignment_cd;
    }

    public void setVertical_alignment_cd(int vertical_alignment_cd) {
        this.vertical_alignment_cd = vertical_alignment_cd;
    }

    public int getCell_type() {
        return cell_type;
    }

    public void setCell_type(int cell_type) {
        this.cell_type = cell_type;
    }

    public short getFill_pattern() {
        return fill_pattern;
    }

    public void setFill_pattern(short fill_pattern) {
        this.fill_pattern = fill_pattern;
    }

    public String getString_value() {
        return string_value;
    }

    public void setString_value(String string_value) {
        this.string_value = string_value;
    }

    public double getDouble_value() {
        return double_value;
    }

    public void setDouble_value(double double_value) {
        this.double_value = double_value;
    }

    public boolean isBoolean_value() {
        return boolean_value;
    }

    public void setBoolean_value(boolean boolean_value) {
        this.boolean_value = boolean_value;
    }

    public byte getError_value() {
        return error_value;
    }

    public void setError_value(byte error_value) {
        this.error_value = error_value;
    }

    public String getFormula_value() {
        return formula_value;
    }

    public void setFormula_value(String formula_value) {
        this.formula_value = formula_value;
    }

    public int getCell_index() {
        return cell_index;
    }

    public void setCell_index(int cell_index) {
        this.cell_index = cell_index;
    }

    public boolean isAppendCellFlag() {
        return appendCellFlag;
    }

    public void setAppendCellFlag(boolean appendCellFlag) {
        this.appendCellFlag = appendCellFlag;
    }

    /**
     * 개발자가 신규로 동적인 Cell을 생성하기 위하여 기본적인 값을 만든다. 여기서 Cell의 색 및 기타 커스텀 데이터를 집어넣는 방향으로 간다.
     * 기타 값들은 기본값을 셋업 해두었음.
     * @param customCellInfo
     * @return
     */
    public CellInfoBean dumyCreateCell(CellInfoBean customCellInfo){
        CellInfoBean cellInfoBean = new CellInfoBean();

        if(null != customCellInfo ){
            String sheetKey = customCellInfo.getSheet_id();
            String rowKey = customCellInfo.getRow_id();
            int cellIndex = customCellInfo.getCell_index();

            cellInfoBean.setCell_id(UUID.randomUUID().toString());
            cellInfoBean.setSheet_id(sheetKey);
            cellInfoBean.setRow_id(rowKey);
            cellInfoBean.setCell_index(cellIndex);
        }


        return cellInfoBean;
    }


}
