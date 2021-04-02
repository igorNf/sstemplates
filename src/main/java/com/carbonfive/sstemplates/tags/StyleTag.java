package com.carbonfive.sstemplates.tags;

import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;

import com.carbonfive.sstemplates.SsTemplateContext;
import com.carbonfive.sstemplates.SsTemplateException;
import com.carbonfive.sstemplates.hssf.StyleData;

/**
 *
 * @author sivoh
 * @version $REVISION
 */
@SuppressWarnings("unchecked")
public class StyleTag extends BaseTag {

    public static final String      STYLE_DATA_KEY    = "HssfStyleTag.style-data-key";

    private static final String[][] ATTRIBUTES        = new String[][] {
            { "align", "center", Integer.toString(HorizontalAlignment.CENTER.getCode()) },
            { "align", "center-selection", Integer.toString(HorizontalAlignment.CENTER_SELECTION.getCode()) },
            { "align", "fill", Integer.toString(HorizontalAlignment.FILL.getCode()) },
            { "align", "general", Integer.toString(HorizontalAlignment.GENERAL.getCode()) },
            { "align", "left", Integer.toString(HorizontalAlignment.LEFT.getCode()) },
            { "align", "right", Integer.toString(HorizontalAlignment.RIGHT.getCode()) },

            { "valign", "bottom", Integer.toString(VerticalAlignment.BOTTOM.getCode()) },
            { "valign", "center", Integer.toString(VerticalAlignment.CENTER.getCode()) },
            { "valign", "justify", Integer.toString(VerticalAlignment.JUSTIFY.getCode()) },
            { "valign", "top", Integer.toString(VerticalAlignment.TOP.getCode()) },

            { "border", "dash-dot", Integer.toString(BorderStyle.DASH_DOT.getCode()) },
            { "border", "dash-dot-dot", Integer.toString(BorderStyle.DASH_DOT_DOT.getCode()) },
            { "border", "dashed", Integer.toString(BorderStyle.DASHED.getCode()) },
            { "border", "dotted", Integer.toString(BorderStyle.DOTTED.getCode()) },
            { "border", "double", Integer.toString(BorderStyle.DOUBLE.getCode()) },
            { "border", "hair", Integer.toString(BorderStyle.HAIR.getCode()) },
            { "border", "medium", Integer.toString(BorderStyle.MEDIUM.getCode()) },
            { "border", "medium-dash-dot", Integer.toString(BorderStyle.MEDIUM_DASH_DOT.getCode()) },
            { "border", "medium-dash-dot-dot", Integer.toString(BorderStyle.MEDIUM_DASH_DOT_DOT.getCode()) },
            { "border", "medium-dashed", Integer.toString(BorderStyle.MEDIUM_DASHED.getCode()) },
            { "border", "none", Integer.toString(BorderStyle.NONE.getCode()) },
            { "border", "slanted-dash-dot", Integer.toString(BorderStyle.SLANTED_DASH_DOT.getCode()) },
            { "border", "thick", Integer.toString(BorderStyle.THICK.getCode()) },
            { "border", "thin", Integer.toString(BorderStyle.THIN.getCode()) },

            { "pattern", "alt-bars", Integer.toString(FillPatternType.ALT_BARS.getCode()) },
            { "pattern", "big-spots", Integer.toString(FillPatternType.BIG_SPOTS.getCode()) },
            { "pattern", "bricks", Integer.toString(FillPatternType.BRICKS.getCode()) },
            { "pattern", "diamonds", Integer.toString(FillPatternType.DIAMONDS.getCode()) },
            { "pattern", "fine-dots", Integer.toString(FillPatternType.FINE_DOTS.getCode()) },
            { "pattern", "no-fill", Integer.toString(FillPatternType.NO_FILL.getCode()) },
            { "pattern", "solid-foreground", Integer.toString(FillPatternType.SOLID_FOREGROUND.getCode()) },
            { "pattern", "sparse-dots", Integer.toString(FillPatternType.SPARSE_DOTS.getCode()) },
            { "pattern", "squares", Integer.toString(FillPatternType.SQUARES.getCode()) },
            { "pattern", "thick-backward-diag", Integer.toString(FillPatternType.THICK_BACKWARD_DIAG.getCode()) },
            { "pattern", "thick-forward-diag", Integer.toString(FillPatternType.THICK_FORWARD_DIAG.getCode()) },
            { "pattern", "thick-horz-bands", Integer.toString(FillPatternType.THICK_HORZ_BANDS.getCode()) },
            { "pattern", "thick-vert-bands", Integer.toString(FillPatternType.THICK_VERT_BANDS.getCode()) },
            { "pattern", "thin-backward-diag", Integer.toString(FillPatternType.THIN_BACKWARD_DIAG.getCode()) },
            { "pattern", "thin-forward-diag", Integer.toString(FillPatternType.THIN_FORWARD_DIAG.getCode()) },
            { "pattern", "thin-horz-bands", Integer.toString(FillPatternType.THIN_HORZ_BANDS.getCode()) },
            { "pattern", "thin-vert-bands", Integer.toString(FillPatternType.THIN_VERT_BANDS.getCode()) },

            { "fontWeight", "normal", Integer.toString(0) },
            { "fontWeight", "bold", Integer.toString(1) },

            { "typeOffset", "none", Integer.toString(Font.SS_NONE) },
            { "typeOffset", "super", Integer.toString(Font.SS_SUPER) },
            { "typeOffset", "sub", Integer.toString(Font.SS_SUB) },

            { "underline", "none", Integer.toString(Font.U_NONE) },
            { "underline", "double", Integer.toString(Font.U_DOUBLE) },
            { "underline", "double-accounting", Integer.toString(Font.U_DOUBLE_ACCOUNTING) },
            { "underline", "single", Integer.toString(Font.U_SINGLE) },
            { "underline", "single-accounting", Integer.toString(Font.U_SINGLE_ACCOUNTING) },

            { "fontColor", "normal", Integer.toString(Font.COLOR_NORMAL) },
            { "fontColor", "red", Integer.toString(Font.COLOR_RED) }, };

    private static HashMap          attributes        = new HashMap();
    static {
        for (int i = 0; i < ATTRIBUTES.length; i++) {
            HashMap propValues = (HashMap) attributes.get(ATTRIBUTES[i][0]);
            if (propValues == null) {
                propValues = new HashMap();
                attributes.put(ATTRIBUTES[i][0], propValues);
            }
            propValues.put(ATTRIBUTES[i][1], new Integer(ATTRIBUTES[i][2]));
        }

        attributes.put("colors", getColorAttributeValues());
    }

    private String                  name              = null;
    private String                  align             = null;
    private String                  borderBottom      = null;
    private String                  borderTop         = null;
    private String                  borderLeft        = null;
    private String                  borderRight       = null;
    private String                  border            = null;
    private String                  bottomBorderColor = null;
    private String                  topBorderColor    = null;
    private String                  leftBorderColor   = null;
    private String                  rightBorderColor  = null;
    private String                  borderColor       = null;
    private String                  dataFormat        = null;
    private String                  background        = null;
    private String                  foreground        = null;
    private String                  fillPattern       = null;
    private String                  hidden            = null;
    private String                  locked            = null;
    private String                  wrapText          = null;
    private String                  indention         = null;
    private String                  rotation          = null;
    private String                  valign            = null;

    private String                  fontName          = null;
    private String                  fontHeight        = null;
    private String                  typeOffset        = null;
    private String                  fontWeight        = null;
    private String                  fontColor         = null;
    private String                  underline         = null;
    private String                  italic            = null;
    private String                  strikeout         = null;

    private String                  columnWidth       = null;
    private String                  rowHeight         = null;
    private String                  autoColumnWidth   = null;

    public void render (SsTemplateContext context) throws SsTemplateException {
        String oldStyle = context.getCurrentStyle();

        StyleData styleData = new StyleData();

        setAlignment(styleData, context);
        setBorderStyles(styleData, context);
        setColors(styleData, context);
        setDataFormat(styleData, context);
        setFillPattern(styleData, context);
        setFlags(styleData, context);
        setIndentionAndRotation(styleData, context);
        setFontInformation(styleData, context);
        getStyleDataAttributes(styleData, context);

        String parsedName = getStyleName(context);
        parsedName = context.addStyleData(parsedName, styleData);

        String newStyle = (oldStyle.length() > 0) ? oldStyle + " " + parsedName : parsedName;

        context.setCurrentStyle(newStyle);
        renderChildren(context);
        context.setCurrentStyle(oldStyle);
    }

    private void getStyleDataAttributes (StyleData styleData, SsTemplateContext context)
            throws SsTemplateException {
        if (columnWidth != null)
            styleData.put(StyleData.COLUMN_WIDTH_ATTRIBUTE, parseExpression(columnWidth,
                    Integer.class, context));
        if (rowHeight != null)
            styleData.put(StyleData.ROW_HEIGHT_ATTRIBUTE, parseExpression(rowHeight,
                    Integer.class, context));
        if (autoColumnWidth != null)
            styleData.put(StyleData.AUTO_COLUMN_WIDTH_ATTRIBUTE, parseExpression(
                    autoColumnWidth, Boolean.class, context));
    }

    private String getStyleName (SsTemplateContext context) throws SsTemplateException {
        String parsedName = null;
        if (name != null) {
            parsedName = (String) parseExpression(name, String.class, context);
            if (parsedName.indexOf(" ") > -1)
                throw new SsTemplateException("Style name cannot contain spaces: " + parsedName);
        } else {
            // parsedName = context.getUniqueStyleName();
        }
        return parsedName;
    }

    private void setFontInformation (StyleData styleData, SsTemplateContext context)
            throws SsTemplateException {

        if (fontName != null)
            styleData.put("fontName", parseExpression(fontName, String.class, context));

        if (fontHeight != null)
            styleData.put("fontHeight", parseExpression(fontHeight, Integer.class, context));

        if (italic != null)
            styleData.put("italic", parseExpression(italic, Boolean.class, context));

        if (strikeout != null)
            styleData.put("strikeout", parseExpression(strikeout, Boolean.class, context));

        if (fontColor != null)
            setColor(styleData, "fontColor", fontColor, (Map) attributes.get("colors"), context);

        if (typeOffset != null)
            styleData.put("typeOffset", new Integer(findShortValueForAttribute("typeOffset",
                    "typeOffset", typeOffset, context)));

        if (fontWeight != null)
            styleData.put("fontWeight", new Integer(findShortValueForAttribute("fontWeight",
                    "fontWeight", fontWeight, context)));

        if (underline != null)
            styleData.put("underline", new Integer(findShortValueForAttribute("underline",
                    "underline", underline, context)));

    }

    private void setAlignment (StyleData styleData, SsTemplateContext context)
            throws SsTemplateException {
        if (align != null)
            styleData.put("align", new Integer(findShortValueForAttribute("align", "align", align,
                    context)));

        if (valign != null)
            styleData.put("valign", new Integer(findShortValueForAttribute("valign", "valign",
                    valign, context)));
    }

    private void setIndentionAndRotation (StyleData styleData, SsTemplateContext context)
            throws SsTemplateException {
        if (indention != null) {
            Integer indent = (Integer) parseExpression(indention, Integer.class, context);
            if (indent.intValue() >= 16)
                throw new SsTemplateException(
                        "Error: indention attribute of style cannot exceed 15");
            styleData.put("indention", indent);
        }

        if (rotation != null)
            styleData.put("rotation", parseExpression(rotation, Integer.class, context));
    }

    private void setFlags (StyleData styleData, SsTemplateContext context)
            throws SsTemplateException {
        if (hidden != null)
            styleData.put("hidden", parseExpression(hidden, Boolean.class, context));

        if (locked != null)
            styleData.put("locked", parseExpression(locked, Boolean.class, context));

        if (wrapText != null)
            styleData.put("wrapText", parseExpression(wrapText, Boolean.class, context));
    }

    private void setFillPattern (StyleData styleData, SsTemplateContext context)
            throws SsTemplateException {
        if (fillPattern != null)
            styleData.put("fillPattern", new Integer(findShortValueForAttribute("fillPattern",
                    "pattern", fillPattern, context)));
    }

    private void setDataFormat (StyleData styleData, SsTemplateContext context)
            throws SsTemplateException {
        if (dataFormat != null) {
            String formatString = (String) parseExpression(dataFormat, String.class, context);
            styleData.put("dataFormat", new Integer(context.getWorkbook().createDataFormat()
                    .getFormat(formatString)));
        }
    }

    private void setColors (StyleData styleData, SsTemplateContext context)
            throws SsTemplateException {
        if (borderColor != null)
            setColor(styleData, "borderColor", borderColor, (Map) attributes.get("colors"), context);

        if (topBorderColor != null)
            setColor(styleData, "topBorderColor", topBorderColor, (Map) attributes.get("colors"), context);

        if (bottomBorderColor != null)
            setColor(styleData, "bottomBorderColor", bottomBorderColor, (Map) attributes.get("colors"), context);

        if (rightBorderColor != null)
            setColor(styleData, "rightBorderColor", rightBorderColor, (Map) attributes.get("colors"), context);

        if (leftBorderColor != null)
            setColor(styleData, "leftBorderColor", leftBorderColor, (Map) attributes.get("colors"), context);

        if (foreground != null)
            setColor(styleData, "foreground", foreground, (Map) attributes.get("colors"), context);

        if (background != null)
            setColor(styleData, "background", background, (Map) attributes.get("colors"), context);
    }

    private void setColor (StyleData styleData, String name, String value, Map colorMap,
            SsTemplateContext context) throws SsTemplateException {
        short[] triplet;
        short index;
        String parsedValue = (String) parseExpression(value, String.class, context);
        if (parsedValue.startsWith("#")){
            triplet = parseColor(parsedValue);
            // Byte[] bytes = {(byte)triplet[0], (byte)triplet[1], (byte)triplet[2]};
            // index = context.getColorIndex(triplet);
            // XSSFColor col = new XSSFColor(bytes, null);
            System.out.println("Кладем значение !!!!!");
            System.out.println(Arrays.toString(triplet));
            styleData.put(name, triplet);
        } else if (colorMap.containsKey(parsedValue)) {
            index = ((IndexedColors) colorMap.get(parsedValue)).getIndex();
            styleData.put(name, new Integer(index));
        }  else {
            throw new SsTemplateException("Can't understand value '" + parsedValue + "' for color '" + name + "'");
        }

        // styleData.put(name, new Integer(index));
        // styleData.put(name, new Integer(context.getColorIndex(triplet)));
    }

    public static short[] parseColor (String value) throws SsTemplateException {
        if (value.length() != 7)
            throw new SsTemplateException("Unable to parse color '" + value + "'");

        short[] triplet = new short[3];

        try {
            triplet[0] = Short.parseShort(value.substring(1, 3), 16);
            triplet[1] = Short.parseShort(value.substring(3, 5), 16);
            triplet[2] = Short.parseShort(value.substring(5, 7), 16);
        } catch (NumberFormatException e) {
            throw new SsTemplateException("Unable to parse color '" + value + "'");
        }
        return triplet;
    }

    private void setBorderStyles (StyleData styleData, SsTemplateContext context)
            throws SsTemplateException {
        if (border != null) {
            styleData.put("border", new Integer(findShortValueForAttribute("border", "border",
                    border, context)));
            if (borderColor == null && topBorderColor == null) topBorderColor = "black";
            if (borderColor == null && bottomBorderColor == null) bottomBorderColor = "black";
            if (borderColor == null && rightBorderColor == null) rightBorderColor = "black";
            if (borderColor == null && leftBorderColor == null) leftBorderColor = "black";
        }

        if (borderTop != null) {
            styleData.put("borderTop", new Integer(findShortValueForAttribute("borderTop",
                    "border", borderTop, context)));
            if (borderColor == null && topBorderColor == null) topBorderColor = "black";
        }

        if (borderBottom != null) {
            styleData.put("borderBottom", new Integer(findShortValueForAttribute("borderBottom",
                    "border", borderBottom, context)));
            if (borderColor == null && bottomBorderColor == null) bottomBorderColor = "black";
        }

        if (borderRight != null) {
            styleData.put("borderRight", new Integer(findShortValueForAttribute("borderRight",
                    "border", borderRight, context)));
            if (borderColor == null && rightBorderColor == null) rightBorderColor = "black";
        }

        if (borderLeft != null) {
            styleData.put("borderLeft", new Integer(findShortValueForAttribute("borderLeft",
                    "border", borderLeft, context)));
            if (borderColor == null && leftBorderColor == null) leftBorderColor = "black";
        }
    }

    private short findShortValueForAttribute (String errorName, String attributeName,
            String attributeValue, SsTemplateContext context) throws SsTemplateException {
        String parsedAttribute = (String) parseExpression(attributeValue, String.class, context);
        Integer value = (Integer) ((HashMap) attributes.get(attributeName)).get(parsedAttribute);
        if (value == null)
            throw new SsTemplateException("Unknown value '" + parsedAttribute + "' for "
                    + errorName + " attribute of style tag");
        short result = value.shortValue();
        return result;
    }

    private static HashMap getColorAttributeValues () {
        HashMap<String, IndexedColors> colors = new HashMap();
        // Class[] colorClasses = HSSFColor.class.getClasses();
        // for (int i = 0; i < colorClasses.length; i++) {
        //     try {
        //         if (HSSFColor.class.isAssignableFrom(colorClasses[i]))
        //         // colors.put(classNameToAttributeValue(colorClasses[i]),new
        //             // Integer(getStaticShortField(colorClasses[i],"index")));
        //         colors.put(classNameToAttributeValue(colorClasses[i]), colorClasses[i].getConstructor().newInstance());
        //     } catch (Exception e) {
        //         // should never happen
        //         log.fine(e.getMessage());
        //     }
        // }

        for (IndexedColors color : IndexedColors.values()) {
            colors.put(color.toString().toLowerCase(), color);
        }

        return colors;
    }

    private static String classNameToAttributeValue (Class clazz) {
        String className = clazz.getName();
        int dollarIndex = className.lastIndexOf("$");
        if (dollarIndex >= 0) className = className.substring(dollarIndex + 1);
        className = className.toLowerCase();
        return className.replace('_', '-');
    }

    public String getAlign () {
        return align;
    }

    public void setAlign (String align) {
        this.align = align;
    }

    public String getBorderBottom () {
        return borderBottom;
    }

    public void setBorderBottom (String borderBottom) {
        this.borderBottom = borderBottom;
    }

    public String getBorderTop () {
        return borderTop;
    }

    public void setBorderTop (String borderTop) {
        this.borderTop = borderTop;
    }

    public String getBorderLeft () {
        return borderLeft;
    }

    public void setBorderLeft (String borderLeft) {
        this.borderLeft = borderLeft;
    }

    public String getBorderRight () {
        return borderRight;
    }

    public void setBorderRight (String borderRight) {
        this.borderRight = borderRight;
    }

    public String getBorder () {
        return border;
    }

    public void setBorder (String border) {
        this.border = border;
    }

    public String getBottomBorderColor () {
        return bottomBorderColor;
    }

    public void setBottomBorderColor (String bottomBorderColor) {
        this.bottomBorderColor = bottomBorderColor;
    }

    public String getTopBorderColor () {
        return topBorderColor;
    }

    public void setTopBorderColor (String topBorderColor) {
        this.topBorderColor = topBorderColor;
    }

    public String getLeftBorderColor () {
        return leftBorderColor;
    }

    public void setLeftBorderColor (String leftBorderColor) {
        this.leftBorderColor = leftBorderColor;
    }

    public String getRightBorderColor () {
        return rightBorderColor;
    }

    public void setRightBorderColor (String rightBorderColor) {
        this.rightBorderColor = rightBorderColor;
    }

    public String getBorderColor () {
        return borderColor;
    }

    public void setBorderColor (String borderColor) {
        this.borderColor = borderColor;
    }

    public String getDataFormat () {
        return dataFormat;
    }

    public void setDataFormat (String dataFormat) {
        this.dataFormat = dataFormat;
    }

    public String getBackground () {
        return background;
    }

    public void setBackground (String background) {
        this.background = background;
    }

    public String getForeground () {
        return foreground;
    }

    public void setForeground (String foreground) {
        this.foreground = foreground;
    }

    public String getFillPattern () {
        return fillPattern;
    }

    public void setFillPattern (String fillPattern) {
        this.fillPattern = fillPattern;
    }

    public String getHidden () {
        return hidden;
    }

    public void setHidden (String hidden) {
        this.hidden = hidden;
    }

    public String getIndention () {
        return indention;
    }

    public void setIndention (String indention) {
        this.indention = indention;
    }

    public String getLocked () {
        return locked;
    }

    public void setLocked (String locked) {
        this.locked = locked;
    }

    public String getRotation () {
        return rotation;
    }

    public void setRotation (String rotation) {
        this.rotation = rotation;
    }

    public String getValign () {
        return valign;
    }

    public void setValign (String valign) {
        this.valign = valign;
    }

    public String getWrapText () {
        return wrapText;
    }

    public void setWrapText (String wrapText) {
        this.wrapText = wrapText;
    }

    public String getFontWeight () {
        return fontWeight;
    }

    public void setFontWeight (String fontWeight) {
        this.fontWeight = fontWeight;
    }

    public String getFontColor () {
        return fontColor;
    }

    public void setFontColor (String fontColor) {
        this.fontColor = fontColor;
    }

    public String getFontHeight () {
        return fontHeight;
    }

    public void setFontHeight (String fontHeight) {
        this.fontHeight = fontHeight;
    }

    public String getFontName () {
        return fontName;
    }

    public void setFontName (String fontName) {
        this.fontName = fontName;
    }

    public String getItalic () {
        return italic;
    }

    public void setItalic (String italic) {
        this.italic = italic;
    }

    public String getStrikeout () {
        return strikeout;
    }

    public void setStrikeout (String strikeout) {
        this.strikeout = strikeout;
    }

    public String getTypeOffset () {
        return typeOffset;
    }

    public void setTypeOffset (String typeOffset) {
        this.typeOffset = typeOffset;
    }

    public String getUnderline () {
        return underline;
    }

    public void setUnderline (String underline) {
        this.underline = underline;
    }

    public String getName () {
        return name;
    }

    public void setName (String name) {
        this.name = name;
    }

    public String getColumnWidth () {
        return columnWidth;
    }

    public void setColumnWidth (String columnWidth) {
        this.columnWidth = columnWidth;
    }

    public String getRowHeight () {
        return rowHeight;
    }

    public void setRowHeight (String rowHeight) {
        this.rowHeight = rowHeight;
    }

    public String getAutoColumnWidth () {
        return autoColumnWidth;
    }

    public void setAutoColumnWidth (String autoColumnWidth) {
        this.autoColumnWidth = autoColumnWidth;
    }
}
