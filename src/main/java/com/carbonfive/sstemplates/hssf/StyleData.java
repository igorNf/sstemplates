package com.carbonfive.sstemplates.hssf;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

import com.carbonfive.sstemplates.SsTemplateContext;
import com.carbonfive.sstemplates.SsTemplateException;


/**
 * @author sivoh
 * @version $REVISION
 */
public class StyleData {
    public static final String      COLUMN_WIDTH_ATTRIBUTE      = "columnWidth";
    public static final String      ROW_HEIGHT_ATTRIBUTE        = "rowHeight";
    public static final String      AUTO_COLUMN_WIDTH_ATTRIBUTE = "autoColumnWidth";

    private HashMap<String, Object> styleData                   = new HashMap<String, Object>();

    public StyleData() {
        // do nothing
    }

    public StyleData(StyleData[] datas) {
        for (int i = 0; i < datas.length; i++)
            datas[i].overideAttributes(styleData);
    }

    public void put (String attribute, Object value) {
        styleData.put(attribute, value);
    }

    public Object get (String attribute) {
        return styleData.get(attribute);
    }

    public boolean containsKey (String attribute) {
        return styleData.containsKey(attribute);
    }

    public void overideAttributes (Map<String, Object> originalStyle) {
        originalStyle.putAll(styleData);
    }

    public void setStyleAttributes (CellStyle style, SsTemplateContext context)
            throws SsTemplateException {
        if (styleData.containsKey("border")) {
            short borderCode = ((Integer) styleData.get("border")).shortValue();
            BorderStyle borderStyle = BorderStyle.valueOf(borderCode);
            style.setBorderTop(borderStyle);
            style.setBorderBottom(borderStyle);
            style.setBorderRight(borderStyle);
            style.setBorderLeft(borderStyle);
        }
        style.setBorderTop(BorderStyle.valueOf(shortFromStyleData("borderTop", style.getBorderTop().getCode())));
        style.setBorderBottom(BorderStyle.valueOf(shortFromStyleData("borderBottom", style.getBorderBottom().getCode())));
        style.setBorderRight(BorderStyle.valueOf(shortFromStyleData("borderRight", style.getBorderRight().getCode())));
        style.setBorderLeft(BorderStyle.valueOf(shortFromStyleData("borderLeft", style.getBorderLeft().getCode())));

        if (styleData.containsKey("borderColor")) {
            short bc = ((Integer) styleData.get("borderColor")).shortValue();
            style.setTopBorderColor(bc);
            style.setBottomBorderColor(bc);
            style.setRightBorderColor(bc);
            style.setLeftBorderColor(bc);
        }

        style.setTopBorderColor(shortFromStyleData("topBorderColor", style.getTopBorderColor()));
        style.setBottomBorderColor(shortFromStyleData("bottomBorderColor", style.getBottomBorderColor()));
        style.setRightBorderColor(shortFromStyleData("rightBorderColor", style.getRightBorderColor()));
        style.setLeftBorderColor(shortFromStyleData("leftBorderColor", style.getLeftBorderColor()));


        System.out.println("============== " + style.getFillForegroundColor());
        System.out.println("============== " + style.getFillBackgroundColor());
        Object foreground = colorFromStyleData("foreground", style.getFillForegroundColor());
        Object background = colorFromStyleData("background", style.getFillBackgroundColor());

        if (foreground instanceof Integer) {
            style.setFillForegroundColor(((Integer) foreground).shortValue());
        } else if (foreground instanceof short[]) {
            short[] f = (short[]) foreground;
            if (style instanceof HSSFCellStyle) {
                HSSFCellStyle hStyle = (HSSFCellStyle) style;
                short s = context.getColorIndex((short[]) foreground);
                hStyle.setFillForegroundColor(s);
            } else {
                XSSFCellStyle xStyle = (XSSFCellStyle) style;
                xStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(f[0], f[1], f[2]), null));
            }
        }
        if (background instanceof Integer) {
            style.setFillForegroundColor(((Integer) background).shortValue());
        } else if (background instanceof short[]) {
            short[] f = (short[]) background;
            if (style instanceof HSSFCellStyle) {
                HSSFCellStyle hStyle = (HSSFCellStyle) style;
                short s = context.getColorIndex((short[]) background);
                hStyle.setFillBackgroundColor(s);
            } else {
                XSSFCellStyle xStyle = (XSSFCellStyle) style;
                xStyle.setFillBackgroundColor(new XSSFColor(new java.awt.Color(f[0], f[1], f[2]), null));
            }
        }

        //style.setFillForegroundColor(shortFromStyleData("foreground", style.getFillForegroundColor()));
        // style.setFillBackgroundColor(shortFromStyleData("background", style.getFillBackgroundColor()));

//        if (!styleData.containsKey("foreground") && (context.getBackgroundColor() != null)) {
//            style.setFillForegroundColor(context.getColorIndex(context.getBackgroundColor()));
//            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
//        }

        style.setDataFormat(shortFromStyleData("dataFormat", style.getDataFormat()));

        style.setFillPattern(FillPatternType.forInt(shortFromStyleData("fillPattern", style.getFillPattern().getCode())));

        style.setHidden(booleanFromStyleData("hidden", style.getHidden()));
        style.setLocked(booleanFromStyleData("locked", style.getLocked()));
        style.setWrapText(booleanFromStyleData("wrapText", style.getWrapText()));

        style.setIndention(shortFromStyleData("indention", style.getIndention()));
        style.setRotation(shortFromStyleData("rotation", style.getRotation()));

        style.setAlignment(HorizontalAlignment.forInt(shortFromStyleData("align", style.getAlignment().getCode())));
        style.setVerticalAlignment(VerticalAlignment.forInt(shortFromStyleData("valign", style.getVerticalAlignment().getCode())));

        Font oldFont = context.getWorkbook().getFontAt(style.getFontIndex());
        String parsedFontName = stringFromStyleData("fontName", oldFont.getFontName());
        boolean parsedItalic = booleanFromStyleData("italic", oldFont.getItalic());
        boolean parsedStrikeout = booleanFromStyleData("strikeout", oldFont.getStrikeout());
        short parsedFontHeight = shortFromStyleData("fontHeight", oldFont.getFontHeight());
        short parsedFontColor = shortFromStyleData("fontColor", oldFont.getColor());
        boolean parsedFontWeight = booleanFromStyleData("fontWeight", oldFont.getBold());
        short parsedTypeOffset = shortFromStyleData("typeOffset", oldFont.getTypeOffset());
        byte parsedUnderline = (byte) shortFromStyleData("underline", oldFont.getUnderline());

        Font font = context.createFont(parsedFontName, parsedFontHeight, parsedFontColor,
                parsedFontWeight, parsedItalic, parsedStrikeout, parsedUnderline, parsedTypeOffset);
        style.setFont(font);
    }

    public Integer getColumnWidth () {
        return (Integer) styleData.get(COLUMN_WIDTH_ATTRIBUTE);
    }

    public Integer getRowHeight () {
        return (Integer) styleData.get(ROW_HEIGHT_ATTRIBUTE);
    }

    public boolean getAutoColumnWidth () {
        Boolean acw = (Boolean) styleData.get(AUTO_COLUMN_WIDTH_ATTRIBUTE);
        return (acw != null) && (acw.booleanValue());
    }

    private short shortFromStyleData (String attribute, short defaultValue) {
        short value = defaultValue;
        if (styleData.containsKey(attribute))
            value = ((Integer) styleData.get(attribute)).shortValue();
        return value;
    }

    private Object colorFromStyleData (String attribute, Object defaultValue) {
        Object value = defaultValue;
        if (styleData.containsKey(attribute))
            value =  styleData.get(attribute);
        return value;
    }


    private boolean booleanFromStyleData (String attribute, boolean defaultValue) {
        boolean value = defaultValue;
        if (styleData.containsKey(attribute))
            value = ((Boolean) styleData.get(attribute)).booleanValue();
        return value;
    }

    private String stringFromStyleData (String attribute, String defaultValue) {
        String value = defaultValue;
        if (styleData.containsKey(attribute)) value = (String) styleData.get(attribute);
        return value;
    }

    public int hashCode () {
        return styleData.hashCode();
    }

    public boolean equals (Object obj) {
        return ((StyleData) obj).styleData.equals(styleData);
    }
}
