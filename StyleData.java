
public class StyleData {
    private short fontSize = 0;
    private String fontName;
    private short bgColorIndex = 0;
    private String textalign;
    private String verticalalign;
    private boolean boldtext = false;
    private boolean bordertop = false;
    private boolean borderbottom = false;
    private boolean borderleft = false;
    private boolean borderright = false;
    private boolean locked = true;

    public StyleData() {
    }

    public void setFontSize(short fontSize) {
        this.fontSize = fontSize;
    }

    public short getFontSize() {
        return this.fontSize;
    }

    public void setFontName(String fontName) {
        this.fontName = fontName;
    }

    public String getFontName() {
        return this.fontName;
    }

    public void setBgColorIndex(short bgColorIndex) {
        this.bgColorIndex = bgColorIndex;
    }

    public short getBgColorIndex() {
        return this.bgColorIndex;
    }

    public void setTextalign(String textalign) {
        this.textalign = textalign;
    }

    public String getTextalign() {
        return this.textalign;
    }

    public void setVerticalalign(String verticalalign) {
        this.verticalalign = verticalalign;
    }

    public String getVerticalalign() {
        return this.verticalalign;
    }

    public void setBoldtext(boolean boldtext) {
        this.boldtext = boldtext;
    }

    public boolean isBoldtext() {
        return this.boldtext;
    }

    public void setBordertop(boolean bordertop) {
        this.bordertop = bordertop;
    }

    public boolean isBordertop() {
        return this.bordertop;
    }

    public void setBorderbottom(boolean borderbottom) {
        this.borderbottom = borderbottom;
    }

    public boolean isBorderbottom() {
        return this.borderbottom;
    }

    public void setBorderleft(boolean borderleft) {
        this.borderleft = borderleft;
    }

    public boolean isBorderleft() {
        return this.borderleft;
    }

    public void setBorderright(boolean borderright) {
        this.borderright = borderright;
    }

    public boolean isBorderright() {
        return this.borderright;
    }

    public void setLocked(boolean locked) {
        this.locked = locked;
    }

    public boolean isLocked() {
        return this.locked;
    }
}