import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.POILogFactory;
import org.apache.poi.util.POILogger;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.model.CommentsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTComment;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.util.Comparator;
import java.util.LinkedList;
import java.util.Queue;

public class XSSFCustomSheetXMLHandler extends DefaultHandler {
    private static final POILogger logger = POILogFactory.getLogger(XSSFCustomSheetXMLHandler.class);
    private StylesTable stylesTable;
    private CommentsTable commentsTable;
    private ReadOnlySharedStringsTable sharedStringsTable;
    private final XSSFCustomSheetXMLHandler.SheetContentsHandler output;
    private boolean vIsOpen;
    private boolean fIsOpen;
    private boolean isIsOpen;
    private boolean hfIsOpen;
    private XSSFCustomSheetXMLHandler.xssfDataType nextDataType;
    private short formatIndex;
    private String formatString;
    private final DataFormatter formatter;
    private int rowNum;
    private String cellRef;
    private boolean formulasNotResults;
    private StringBuffer value;
    private StringBuffer formula;
    private XSSFCellStyle xssfCellStyle;
    private StringBuffer headerFooter;
    private Queue<CellReference> commentCellRefs;
    private static final Comparator<CellReference> cellRefComparator = new Comparator<CellReference>() {
        public int compare(CellReference o1, CellReference o2) {
            int result = this.compare(o1.getRow(), o2.getRow());
            if (result == 0) {
                result = this.compare(o1.getCol(), o2.getCol());
            }

            return result;
        }

        public int compare(int x, int y) {
            return x < y ? -1 : (x == y ? 0 : 1);
        }
    };

    public XSSFCustomSheetXMLHandler(StylesTable styles, CommentsTable comments, ReadOnlySharedStringsTable strings, XSSFCustomSheetXMLHandler.SheetContentsHandler sheetContentsHandler, DataFormatter dataFormatter, boolean formulasNotResults) {
        this.value = new StringBuffer();
        this.formula = new StringBuffer();
        this.headerFooter = new StringBuffer();
        this.stylesTable = styles;
        this.commentsTable = comments;
        this.sharedStringsTable = strings;
        this.output = sheetContentsHandler;
        this.formulasNotResults = formulasNotResults;
        this.nextDataType = XSSFCustomSheetXMLHandler.xssfDataType.NUMBER;
        this.formatter = dataFormatter;
        this.init();
    }

    public XSSFCustomSheetXMLHandler(StylesTable styles, ReadOnlySharedStringsTable strings, XSSFCustomSheetXMLHandler.SheetContentsHandler sheetContentsHandler, DataFormatter dataFormatter, boolean formulasNotResults) {
        this(styles, (CommentsTable)null, strings, sheetContentsHandler, dataFormatter, formulasNotResults);
    }

    public XSSFCustomSheetXMLHandler(StylesTable styles, ReadOnlySharedStringsTable strings, XSSFCustomSheetXMLHandler.SheetContentsHandler sheetContentsHandler, boolean formulasNotResults) {
        this(styles, strings, sheetContentsHandler, new DataFormatter(), formulasNotResults);
    }

    private void init() {
        if (this.commentsTable != null) {
            this.commentCellRefs = new LinkedList();
            CTComment[] arr$ = this.commentsTable.getCTComments().getCommentList().getCommentArray();
            int len$ = arr$.length;

            for(int i$ = 0; i$ < len$; ++i$) {
                CTComment comment = arr$[i$];
                this.commentCellRefs.add(new CellReference(comment.getRef()));
            }
        }

    }

    private boolean isTextTag(String name) {
        if ("v".equals(name)) {
            return true;
        } else if ("inlineStr".equals(name)) {
            return true;
        } else {
            return "t".equals(name) && this.isIsOpen;
        }
    }

    public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
        if (this.isTextTag(name)) {
            this.vIsOpen = true;
            this.value.setLength(0);
        } else if ("is".equals(name)) {
            this.isIsOpen = true;
        } else {
            String cellType;
            String cellStyleStr;
            if ("f".equals(name)) {
                this.formula.setLength(0);
                if (this.nextDataType == XSSFCustomSheetXMLHandler.xssfDataType.NUMBER) {
                    this.nextDataType = XSSFCustomSheetXMLHandler.xssfDataType.FORMULA;
                }

                cellType = attributes.getValue("t");
                if (cellType != null && cellType.equals("shared")) {
                    cellStyleStr = attributes.getValue("ref");
                    String si = attributes.getValue("si");
                    if (cellStyleStr != null) {
                        this.fIsOpen = true;
                    } else if (this.formulasNotResults) {
                        logger.log(5, "shared formulas not yet supported!");
                    }
                } else {
                    this.fIsOpen = true;
                }
            } else if (!"oddHeader".equals(name) && !"evenHeader".equals(name) && !"firstHeader".equals(name) && !"firstFooter".equals(name) && !"oddFooter".equals(name) && !"evenFooter".equals(name)) {
                if ("row".equals(name)) {
                    this.rowNum = Integer.parseInt(attributes.getValue("r")) - 1;
                    this.output.startRow(this.rowNum);
                } else if ("c".equals(name)) {
                    this.nextDataType = XSSFCustomSheetXMLHandler.xssfDataType.NUMBER;
                    this.formatIndex = -1;
                    this.formatString = null;
                    this.cellRef = attributes.getValue("r");
                    cellType = attributes.getValue("t");
                    cellStyleStr = attributes.getValue("s");
                    if (stylesTable != null) {
                        if (cellStyleStr != null) {
                            int styleIndex = Integer.parseInt(cellStyleStr);
                            this.xssfCellStyle = stylesTable.getStyleAt(styleIndex);
                        } else if (stylesTable.getNumCellStyles() > 0) {
                            this.xssfCellStyle = stylesTable.getStyleAt(0);
                        }
                    }


                    if ("b".equals(cellType)) {
                        this.nextDataType = XSSFCustomSheetXMLHandler.xssfDataType.BOOLEAN;
                    } else if ("e".equals(cellType)) {
                        this.nextDataType = XSSFCustomSheetXMLHandler.xssfDataType.ERROR;
                    } else if ("inlineStr".equals(cellType)) {
                        this.nextDataType = XSSFCustomSheetXMLHandler.xssfDataType.INLINE_STRING;
                    } else if ("s".equals(cellType)) {
                        this.nextDataType = XSSFCustomSheetXMLHandler.xssfDataType.SST_STRING;
                    } else if ("str".equals(cellType)) {
                        this.nextDataType = XSSFCustomSheetXMLHandler.xssfDataType.FORMULA;
                    } else {
                        XSSFCellStyle style = null;
                        if (cellStyleStr != null) {
                            int styleIndex = Integer.parseInt(cellStyleStr);
                            style = this.stylesTable.getStyleAt(styleIndex);
                        } else if (this.stylesTable.getNumCellStyles() > 0) {
                            style = this.stylesTable.getStyleAt(0);
                        }

                        if (style != null) {
                            this.formatIndex = style.getDataFormat();
                            this.formatString = style.getDataFormatString();
                            if (this.formatString == null) {
                                this.formatString = BuiltinFormats.getBuiltinFormat(this.formatIndex);
                            }
                        }
                    }
                }
            } else {
                this.hfIsOpen = true;
                this.headerFooter.setLength(0);
            }
        }

    }

    public void endElement(String uri, String localName, String name) throws SAXException {
        String thisStr = null;
        if (this.isTextTag(name)) {
            this.vIsOpen = false;
            switch(this.nextDataType) {
                case BOOLEAN:
                    char first = this.value.charAt(0);
                    thisStr = first == '0' ? "FALSE" : "TRUE";
                    break;
                case ERROR:
                    thisStr = "ERROR:" + this.value.toString();
                    break;
                case FORMULA:
                    if (this.formulasNotResults) {
                        thisStr = this.formula.toString();
                    } else {
                        String fv = this.value.toString();
                        if (this.formatString != null) {
                            try {
                                double d = Double.parseDouble(fv);
                                thisStr = this.formatter.formatRawCellContents(d, this.formatIndex, this.formatString);
                            } catch (NumberFormatException var11) {
                                thisStr = fv;
                            }
                        } else {
                            thisStr = fv;
                        }
                    }
                    break;
                case INLINE_STRING:
                    XSSFRichTextString rtsi = new XSSFRichTextString(this.value.toString());
                    thisStr = rtsi.toString();
                    break;
                case SST_STRING:
                    String sstIndex = this.value.toString();

                    try {
                        int idx = Integer.parseInt(sstIndex);
                        XSSFRichTextString rtss = new XSSFRichTextString(this.sharedStringsTable.getEntryAt(idx));
                        thisStr = rtss.toString();
                    } catch (NumberFormatException var10) {
                        logger.log(7, "Failed to parse SST index '" + sstIndex, var10);
                    }
                    break;
                case NUMBER:
                    String n = this.value.toString();
                    if (this.formatString != null) {
                        thisStr = this.formatter.formatRawCellContents(Double.parseDouble(n), this.formatIndex, this.formatString);
                    } else {
                        thisStr = n;
                    }
                    break;
                default:
                    thisStr = "(TODO: Unexpected type: " + this.nextDataType + ")";
            }

            this.checkForEmptyCellComments(XSSFCustomSheetXMLHandler.EmptyCellCommentsCheckType.CELL);
            XSSFComment comment = this.commentsTable != null ? this.commentsTable.findCellComment(this.cellRef) : null;
            this.output.cell(this.cellRef, thisStr, comment,this.xssfCellStyle);
        } else if ("f".equals(name)) {
            this.fIsOpen = false;
        } else if ("is".equals(name)) {
            this.isIsOpen = false;
        } else if ("row".equals(name)) {
            this.checkForEmptyCellComments(XSSFCustomSheetXMLHandler.EmptyCellCommentsCheckType.END_OF_ROW);
            this.output.endRow(this.rowNum);
        } else if ("sheetData".equals(name)) {
            this.checkForEmptyCellComments(XSSFCustomSheetXMLHandler.EmptyCellCommentsCheckType.END_OF_SHEET_DATA);
        } else if (!"oddHeader".equals(name) && !"evenHeader".equals(name) && !"firstHeader".equals(name)) {
            if ("oddFooter".equals(name) || "evenFooter".equals(name) || "firstFooter".equals(name)) {
                this.hfIsOpen = false;
                this.output.headerFooter(this.headerFooter.toString(), false, name);
            }
        } else {
            this.hfIsOpen = false;
            this.output.headerFooter(this.headerFooter.toString(), true, name);
        }

    }

    public void characters(char[] ch, int start, int length) throws SAXException {
        if (this.vIsOpen) {
            this.value.append(ch, start, length);
        }

        if (this.fIsOpen) {
            this.formula.append(ch, start, length);
        }

        if (this.hfIsOpen) {
            this.headerFooter.append(ch, start, length);
        }

    }

    private void checkForEmptyCellComments(XSSFCustomSheetXMLHandler.EmptyCellCommentsCheckType type) {
        if (this.commentCellRefs != null && !this.commentCellRefs.isEmpty()) {
            if (type == XSSFCustomSheetXMLHandler.EmptyCellCommentsCheckType.END_OF_SHEET_DATA) {
                while(!this.commentCellRefs.isEmpty()) {
                    this.outputEmptyCellComment((CellReference)this.commentCellRefs.remove());
                }

                return;
            }

            if (this.cellRef == null) {
                if (type == XSSFCustomSheetXMLHandler.EmptyCellCommentsCheckType.END_OF_ROW) {
                    while(!this.commentCellRefs.isEmpty()) {
                        if (((CellReference)this.commentCellRefs.peek()).getRow() != this.rowNum) {
                            return;
                        }

                        this.outputEmptyCellComment((CellReference)this.commentCellRefs.remove());
                    }

                    return;
                }

                throw new IllegalStateException("Cell ref should be null only if there are only empty cells in the row; rowNum: " + this.rowNum);
            }

            CellReference nextCommentCellRef;
            do {
                CellReference cellRef = new CellReference(this.cellRef);
                CellReference peekCellRef = (CellReference)this.commentCellRefs.peek();
                if (type == XSSFCustomSheetXMLHandler.EmptyCellCommentsCheckType.CELL && cellRef.equals(peekCellRef)) {
                    this.commentCellRefs.remove();
                    return;
                }

                int comparison = cellRefComparator.compare(peekCellRef, cellRef);
                if (comparison > 0 && type == XSSFCustomSheetXMLHandler.EmptyCellCommentsCheckType.END_OF_ROW && peekCellRef.getRow() <= this.rowNum) {
                    nextCommentCellRef = (CellReference)this.commentCellRefs.remove();
                    this.outputEmptyCellComment(nextCommentCellRef);
                } else if (comparison < 0 && type == XSSFCustomSheetXMLHandler.EmptyCellCommentsCheckType.CELL && peekCellRef.getRow() <= this.rowNum) {
                    nextCommentCellRef = (CellReference)this.commentCellRefs.remove();
                    this.outputEmptyCellComment(nextCommentCellRef);
                } else {
                    nextCommentCellRef = null;
                }
            } while(nextCommentCellRef != null && !this.commentCellRefs.isEmpty());
        }

    }

    private void outputEmptyCellComment(CellReference cellRef) {
        String cellRefString = cellRef.formatAsString();
        XSSFComment comment = this.commentsTable.findCellComment(cellRefString);
        this.output.cell(cellRefString, (String)null, comment,this.xssfCellStyle);
    }

    public interface SheetContentsHandler {
        void startRow(int var1);

        void endRow(int var1);

        void cell(String var1, String var2, XSSFComment var3,XSSFCellStyle cellStyle);

        void headerFooter(String var1, boolean var2, String var3);
    }

    private static enum EmptyCellCommentsCheckType {
        CELL,
        END_OF_ROW,
        END_OF_SHEET_DATA;

        private EmptyCellCommentsCheckType() {
        }
    }

    static enum xssfDataType {
        BOOLEAN,
        ERROR,
        FORMULA,
        INLINE_STRING,
        SST_STRING,
        NUMBER;

        private xssfDataType() {
        }
    }
}
