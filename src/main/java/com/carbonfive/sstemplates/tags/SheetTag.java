package com.carbonfive.sstemplates.tags;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.carbonfive.sstemplates.SsTemplateContext;
import com.carbonfive.sstemplates.SsTemplateException;

/**
 *
 * @author sivoh
 * @version $REVISION
 */
public class SheetTag extends BaseTag {

    private String name             = null;
    private String style            = null;
    private String marginTop        = null;
    private String marginBottom     = null;
    private String marginLeft       = null;
    private String marginRight      = null;
    private String marginHeader     = null;
    private String marginFooter     = null;
    private String repeatRowFrom    = null;
    private String repeatRowTo      = null;
    private String repeatColumnFrom = null;
    private String repeatColumnTo   = null;
    private String printRowFrom     = null;
    private String printRowTo       = null;
    private String printColumnFrom  = null;
    private String printColumnTo    = null;
    private String fitWidth         = null;
    private String fitHeight        = null;
    private String landscape        = null;
    private String headerLeft       = null;
    private String headerCenter     = null;
    private String headerRight      = null;
    private String footerLeft       = null;
    private String footerCenter     = null;
    private String footerRight      = null;
    private String zoom             = null;
    private String firstPageBreak   = null;
    private String nextPageBreak    = null;

    private String paperSize        = null;

    public void render (SsTemplateContext context) throws SsTemplateException {

        if (name == null)
            throw new SsTemplateException("Parse Error, sheet tag must have name attribute");

        String parsedName = (String) parseExpression(name, String.class, context);
        Workbook workbook = context.getWorkbook();
        Sheet sheet = workbook.getSheet(parsedName);
        if (sheet == null) sheet = workbook.createSheet(parsedName);

        if (marginTop != null)
            sheet.setMargin(Sheet.TopMargin, ((Double) parseExpression(marginTop, Double.class,
                    context)).doubleValue());
        if (marginBottom != null)
            sheet.setMargin(Sheet.BottomMargin, ((Double) parseExpression(marginBottom,
                    Double.class, context)).doubleValue());
        if (marginLeft != null)
            sheet.setMargin(Sheet.LeftMargin, ((Double) parseExpression(marginLeft,
                    Double.class, context)).doubleValue());
        if (marginRight != null)
            sheet.setMargin(Sheet.RightMargin, ((Double) parseExpression(marginRight,
                    Double.class, context)).doubleValue());
        if (marginHeader != null)
            sheet.getPrintSetup().setHeaderMargin(
                    ((Double) parseExpression(marginHeader, Double.class, context)).doubleValue());
        if (marginFooter != null)
            sheet.getPrintSetup().setFooterMargin(
                    ((Double) parseExpression(marginFooter, Double.class, context)).doubleValue());

        if (((repeatRowFrom != null) && (repeatRowTo != null))
                || ((repeatColumnFrom != null) && (repeatColumnTo != null))) {
            int rowFrom = -1, rowTo = -1;
            int columnFrom = -1, columnTo = -1;

            if ((repeatRowFrom != null) && (repeatRowTo != null)) {
                rowFrom = ((Integer) parseExpression(repeatRowFrom, Integer.class, context))
                        .intValue();
                rowTo = ((Integer) parseExpression(repeatRowTo, Integer.class, context)).intValue();
            }

            if ((repeatColumnFrom != null) && (repeatColumnTo != null)) {
                columnFrom = ((Integer) parseExpression(repeatColumnFrom, Integer.class, context)).intValue();
                columnTo = ((Integer) parseExpression(repeatColumnTo, Integer.class, context)).intValue();
            }

            // workbook.setRepeatingRowsAndColumns(workbook.getSheetIndex(parsedName), columnFrom, columnTo, rowFrom, rowTo);
        }

        // not sure what this really does - seems necessary, though
        if ((fitWidth != null) || (fitHeight != null)) sheet.setAutobreaks(true);

        if (paperSize != null)
            sheet.getPrintSetup().setPaperSize(
                    ((Short) parseExpression(paperSize, Short.class, context)).shortValue());

        if (fitWidth != null)
            sheet.getPrintSetup().setFitWidth(
                    ((Short) parseExpression(fitWidth, Short.class, context)).shortValue());
        if (fitHeight != null)
            sheet.getPrintSetup().setFitHeight(
                    ((Short) parseExpression(fitHeight, Short.class, context)).shortValue());

        if (landscape != null)
            sheet.getPrintSetup().setLandscape(((Boolean) parseExpression(landscape, Boolean.class, context)).booleanValue());

        if (headerLeft != null)
            sheet.getHeader().setLeft((String) parseExpression(headerLeft, String.class, context));

        if (headerCenter != null)
            sheet.getHeader().setCenter((String) parseExpression(headerCenter, String.class, context));

        if (headerRight != null)
            sheet.getHeader().setRight((String) parseExpression(headerRight, String.class, context));

        if (footerLeft != null)
            sheet.getFooter().setLeft((String) parseExpression(footerLeft, String.class, context));

        if (footerCenter != null)
            sheet.getFooter().setCenter((String) parseExpression(footerCenter, String.class, context));

        if (footerRight != null)
            sheet.getFooter().setRight((String) parseExpression(footerRight, String.class, context));

        if (zoom != null)
            sheet.setZoom(((Integer) parseExpression(zoom, Integer.class, context)).intValue());

        if (context.getBackgroundColor() != null) {
            CellStyle cellStyle = sheet.createRow(0).createCell((short) 0).getCellStyle();
            cellStyle.setFillForegroundColor(context.getColorIndex(context.getBackgroundColor()));
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }

        if ((firstPageBreak != null) && (nextPageBreak != null)) {
            context.setPageBreaks(parseInt(firstPageBreak, context), parseInt(nextPageBreak, context));
        }

        String oldStyleName = context.getCurrentStyle();
        context.setCurrentStyle(parseCurrentStyle(oldStyleName, context));
        context.setRowIndex(0);

        context.setSheet(sheet);
        renderChildren(context);
        context.setSheet(null);

        // must be done at the end, because we're computing the max row and
        // column values
        if ((printRowFrom != null) || (printRowTo != null) || (printColumnFrom != null)
                || (printColumnTo != null)) {
            int rowFrom = 0;
            int rowTo = context.getMaxRowIndex();
            int columnFrom = 0;
            int columnTo = context.getMaxColumnIndex();

            if (printRowFrom != null)
                rowFrom = ((Integer) parseExpression(printRowFrom, Integer.class, context))
                        .intValue();

            if (printRowTo != null)
                rowTo = ((Integer) parseExpression(printRowTo, Integer.class, context)).intValue();

            if (printColumnFrom != null)
                columnFrom = ((Integer) parseExpression(printColumnFrom, Integer.class, context))
                        .intValue();

            if (printColumnTo != null)
                columnTo = ((Integer) parseExpression(printColumnTo, Integer.class, context))
                        .intValue();

            workbook.setPrintArea(workbook.getSheetIndex(parsedName), columnFrom, columnTo,
                    rowFrom, rowTo);
        }

        context.setCurrentStyle(oldStyleName);
    }

    private String parseCurrentStyle (String oldStyleName, SsTemplateContext context)
            throws SsTemplateException {

        String styleName = oldStyleName;
        if (getStyle() != null) {
            String parsedStyle = (String) parseExpression(getStyle(), String.class, context);
            styleName = (styleName.length() > 0) ? styleName + " " + parsedStyle : parsedStyle;
        }
        return styleName;
    }

    public String getName () {

        return name;
    }

    public void setName (String name) {

        this.name = name;
    }

    public String getStyle () {

        return style;
    }

    public void setStyle (String style) {

        this.style = style;
    }

    public String getMarginTop () {

        return marginTop;
    }

    public void setMarginTop (String marginTop) {

        this.marginTop = marginTop;
    }

    public String getMarginBottom () {

        return marginBottom;
    }

    public void setMarginBottom (String marginBottom) {

        this.marginBottom = marginBottom;
    }

    public String getMarginLeft () {

        return marginLeft;
    }

    public void setMarginLeft (String marginLeft) {

        this.marginLeft = marginLeft;
    }

    public String getMarginRight () {

        return marginRight;
    }

    public void setMarginRight (String marginRight) {

        this.marginRight = marginRight;
    }

    public String getMarginHeader () {

        return marginHeader;
    }

    public void setMarginHeader (String marginHeader) {

        this.marginHeader = marginHeader;
    }

    public String getMarginFooter () {

        return marginFooter;
    }

    public void setMarginFooter (String marginFooter) {

        this.marginFooter = marginFooter;
    }

    public String getFitWidth () {

        return fitWidth;
    }

    public void setFitWidth (String fitWidth) {

        this.fitWidth = fitWidth;
    }

    public String getFitHeight () {

        return fitHeight;
    }

    public void setFitHeight (String fitHeight) {

        this.fitHeight = fitHeight;
    }

    public String getLandscape () {

        return landscape;
    }

    public void setLandscape (String landscape) {

        this.landscape = landscape;
    }

    public String getRepeatRowFrom () {

        return repeatRowFrom;
    }

    public void setRepeatRowFrom (String repeatRowFrom) {

        this.repeatRowFrom = repeatRowFrom;
    }

    public String getRepeatRowTo () {

        return repeatRowTo;
    }

    public void setRepeatRowTo (String repeatRowTo) {

        this.repeatRowTo = repeatRowTo;
    }

    public String getRepeatColumnFrom () {

        return repeatColumnFrom;
    }

    public void setRepeatColumnFrom (String repeatColumnFrom) {

        this.repeatColumnFrom = repeatColumnFrom;
    }

    public String getRepeatColumnTo () {

        return repeatColumnTo;
    }

    public void setRepeatColumnTo (String repeatColumnTo) {

        this.repeatColumnTo = repeatColumnTo;
    }

    public String getHeaderLeft () {

        return headerLeft;
    }

    public void setHeaderLeft (String headerLeft) {

        this.headerLeft = headerLeft;
    }

    public String getHeaderCenter () {

        return headerCenter;
    }

    public void setHeaderCenter (String headerCenter) {

        this.headerCenter = headerCenter;
    }

    public String getHeaderRight () {

        return headerRight;
    }

    public void setHeaderRight (String headerRight) {

        this.headerRight = headerRight;
    }

    public String getFooterLeft () {

        return footerLeft;
    }

    public void setFooterLeft (String footerLeft) {

        this.footerLeft = footerLeft;
    }

    public String getFooterCenter () {

        return footerCenter;
    }

    public void setFooterCenter (String footerCenter) {

        this.footerCenter = footerCenter;
    }

    public String getFooterRight () {

        return footerRight;
    }

    public void setFooterRight (String footerRight) {

        this.footerRight = footerRight;
    }

    public String getZoom () {

        return zoom;
    }

    public void setZoom (String zoom) {

        this.zoom = zoom;
    }

    public String getPrintRowFrom () {

        return printRowFrom;
    }

    public void setPrintRowFrom (String printRowFrom) {

        this.printRowFrom = printRowFrom;
    }

    public String getPrintRowTo () {

        return printRowTo;
    }

    public void setPrintRowTo (String printRowTo) {

        this.printRowTo = printRowTo;
    }

    public String getPrintColumnFrom () {

        return printColumnFrom;
    }

    public void setPrintColumnFrom (String printColumnFrom) {

        this.printColumnFrom = printColumnFrom;
    }

    public String getPrintColumnTo () {

        return printColumnTo;
    }

    public void setPrintColumnTo (String printColumnTo) {

        this.printColumnTo = printColumnTo;
    }

    public String getFirstPageBreak () {

        return firstPageBreak;
    }

    public void setFirstPageBreak (String firstPageBreak) {

        this.firstPageBreak = firstPageBreak;
    }

    public String getNextPageBreak () {

        return nextPageBreak;
    }

    public void setNextPageBreak (String nextPageBreak) {

        this.nextPageBreak = nextPageBreak;
    }

    public String getPaperSize () {

        return this.paperSize;
    }

    public void setPaperSize (String newPaperSize) {

        this.paperSize = newPaperSize;
    }
}
