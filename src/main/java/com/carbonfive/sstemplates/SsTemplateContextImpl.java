package com.carbonfive.sstemplates;

import java.lang.reflect.Method;
import java.util.Arrays;
import java.util.Collection;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;
import java.util.StringTokenizer;

import org.apache.commons.lang3.builder.EqualsBuilder;
import org.apache.commons.lang3.builder.HashCodeBuilder;

import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import com.carbonfive.sstemplates.hssf.CellAccumulator;
import com.carbonfive.sstemplates.hssf.StyleData;


/**
 * This class acts as an EL VariableResolver, but does not support the
 * pageContext implicit object.
 *
 * @author sivoh
 * @version $REVISION
 */
@SuppressWarnings({"unchecked", "deprecation"})
public class SsTemplateContextImpl implements SsTemplateContext {

    private static final String UNNAMED_STYLE_PREFIX  = "!!!UNNAMED";

    private Map<String, Object> pageScope             = new HashMap<String, Object>();
    private Map                 fontCache             = new HashMap();
    private Map                 styleDataCache        = new HashMap();
    private Map                 styleDataInverseCache = new HashMap();
    private Map                 styleCache            = new HashMap();
    private Map                 accumulatorCache      = new HashMap();
    private Map                 resetAccumulatorCache = new RemoveOnGetMap(accumulatorCache);
    private Map                 functions             = new HashMap();
    private Map                 customValues          = new HashMap();
    private Workbook            workbook              = null;
    private Sheet               sheet                 = null;
    private Row                 row                   = null;
    private String              currentStyle          = "";
    private int                 rowIndex              = 0;
    private int                 columnIndex           = 0;
    private int                 maxRowIndex           = Integer.MIN_VALUE;
    private int                 maxColumnIndex        = Integer.MIN_VALUE;
    private int                 styleCount            = 0;
    private int                 firstPageBreak        = 0;
    private int                 nextPageBreak         = 0;

    private short[]             backgroundColor       = null;

    public SsTemplateContextImpl(Map<String, Object> context) {
        initStyles();
        pageScope = new HashMap<String, Object>(context);
    }

    private void initStyles () {
        StyleData noTopBorder = new StyleData();
        noTopBorder.put("borderTop", new Integer(BorderStyle.NONE.getCode()));
        addStyleData("_noTopBorder", noTopBorder);

        StyleData noBottomBorder = new StyleData();
        noBottomBorder.put("borderBottom", new Integer(BorderStyle.NONE.getCode()));
        addStyleData("_noBottomBorder", noBottomBorder);
    }

    public Object setPageVariable (String key, Object value) {
        return pageScope.put(key, value);
    }

    public void unsetPageVariable (String key, Object oldValue) {
        if (oldValue != null) pageScope.put(key, oldValue);
        else pageScope.remove(key);
    }

    public Object getPageVariable (String key) {
        return pageScope.get(key);
    }

    public Object resolveVariable (String name) {
        if ("pageScope".equals(name)) return pageScope;

        if ("accumulator".equals(name)) return accumulatorCache;

        if ("resetAccumulator".equals(name)) return resetAccumulatorCache;

        // otherwise, try to find the name in page, request, session, then
        // application scope
        if (pageScope.containsKey(name)) return pageScope.get(name);

        return null;
    }

    public Font createFont (String name, short fontHeight, short color, boolean bold,
            boolean italic, boolean strikeout, byte underline, short typeOffset) {
        FontKey fontKey = new FontKey(name, fontHeight, color, bold, italic, strikeout,
                underline, typeOffset);

        Font font = (Font) fontCache.get(fontKey);
        if (font == null) {
            font = workbook.createFont();
            fontKey.setFontProperties(font);
            fontCache.put(fontKey, font);
        }
        return font;
    }

    public void incrementCellIndex () {
        setColumnIndex(columnIndex + 1);
        for (CellRangeAddress region = getRegionForCurrentLocation(); region != null; region = getRegionForCurrentLocation()) {
            if ((columnIndex != region.getFirstColumn()) || (rowIndex != region.getLastRow()))
                setColumnIndex(region.getLastColumn() + 1);
        }
    }

    public void incrementRowIndex () {
        setRowIndex(rowIndex + 1);
        setColumnIndex(-1);
        incrementCellIndex();
    }

    private static boolean _cellRangeAddressContains (CellRangeAddress range, int row, short column) {

        return (range.getFirstRow() <= row) && (row <= range.getLastRow())
                && (range.getFirstColumn() <= column) && (column <= range.getLastColumn());
    }


    public CellRangeAddress getRegionForCurrentLocation () {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            if (_cellRangeAddressContains(sheet.getMergedRegion(i), rowIndex, (short) columnIndex))
                return sheet.getMergedRegion(i);

            // if ( sheet.getMergedRegionAt(i).contains( rowIndex, (short)
            // columnIndex ) )
            // return sheet.getMergedRegionAt(i);
        }
        return null;
    }

    private String getUniqueStyleName () {
        return UNNAMED_STYLE_PREFIX + (styleCount++);
    }

    public String addStyleData (String name, StyleData data) {
        if (name == null) {
            String previous = (String) styleDataInverseCache.get(data);
            if (previous != null) return previous;
            name = getUniqueStyleName();
        }

        styleDataCache.put(name, data);
        styleDataInverseCache.put(data, name);
        return name;
    }

    public String addStyleData (String name, Map data) {
        if (name == null) {
            String previous = (String) styleDataInverseCache.get(data);
            if (previous != null) return previous;
            name = getUniqueStyleName();
        }

        styleDataCache.put(name, data);
        styleDataInverseCache.put(data, name);
        return name;
    }

    public CellStyle getNamedStyle (String name) throws SsTemplateException {
        return getCachedStyleFromName(name).style;
    }

    public StyleData getNamedStyleData (String name) throws SsTemplateException {
        return getCachedStyleFromName(name).styleData;
    }

    private CachedStyle getCachedStyleFromName (String name) throws SsTemplateException {
        CachedStyle cachedStyle = (CachedStyle) styleCache.get(name);
        if (cachedStyle == null) {
            StringTokenizer st = new StringTokenizer(name, " ", false);
            StyleData[] datas = new StyleData[st.countTokens()];
            for (int i = 0; i < datas.length; i++) {
                String token = st.nextToken();
                datas[i] = (StyleData) styleDataCache.get(token);
                if (datas[i] == null)
                    throw new SsTemplateException("Error retrieving undefined style " + token);
            }

            cachedStyle = new CachedStyle(new StyleData(datas), workbook.createCellStyle());
            cachedStyle.styleData.setStyleAttributes(cachedStyle.style, this);
            styleCache.put(name, cachedStyle);
        }
        return cachedStyle;
    }

    public boolean hasCachedStyleData (String name) {
        return styleDataCache.containsKey(name);
    }

    public void setPageBreaks (int firstPageBreak, int nextPageBreak) {
        this.firstPageBreak = firstPageBreak;
        this.nextPageBreak = nextPageBreak;
    }

    public Workbook getWorkbook () {
        return workbook;
    }

    public void setWorkbook (Workbook workbook) {
        this.workbook = workbook;
    }

    public Sheet getSheet () {
        return sheet;
    }

    public void setSheet (Sheet sheet) {
        this.sheet = sheet;
    }

    public Row getRow () {
        return row;
    }

    public void setRow (Row row) {
        this.row = row;
    }

    public int getRowIndex () {
        return rowIndex;
    }

    public void setRowIndex (int rowIndex) {
        if (rowIndex > maxRowIndex) maxRowIndex = rowIndex;
        this.rowIndex = rowIndex;
    }

    public int getColumnIndex () {
        return columnIndex;
    }

    public void setColumnIndex (int columnIndex) {
        if (columnIndex > maxColumnIndex) maxColumnIndex = columnIndex;
        this.columnIndex = columnIndex;
    }

    public String getCurrentStyle () {
        return currentStyle;
    }

    public void setCurrentStyle (String currentStyle) {
        this.currentStyle = currentStyle;
    }

    public CellAccumulator getNamedAccumulator (String name) {
        CellAccumulator acc = (CellAccumulator) accumulatorCache.get(name);

        if (acc == null) {
            acc = new CellAccumulator();
            accumulatorCache.put(name, acc);
            // resetAccumulatorCache.put(name, acc);
        }

        return acc;
    }

    public void registerMethod (String name, Method m) {
        functions.put(name, m);
    }

    // no prefix support
    public Method resolveFunction (String prefix, String name) {
        return (Method) functions.get(name);
    }

    public Object getCustomValue (Object key) {
        return customValues.get(key);
    }

    public void setCustomValue (Object key, Object value) {
        customValues.put(key, value);
    }

    public static final short firstColorIndex   = 0xa;
    public static final short lastColorIndex    = 0x40;
    private short             currentColorIndex = firstColorIndex;
    private Map               colorMap          = new HashMap();
    private HSSFPalette       palette;

    public short getColorIndex (short[] triplet) throws SsTemplateException {
        Color color = new Color(triplet);
        Short index = (Short) colorMap.get(color);

        if (index == null) {
            if (currentColorIndex > lastColorIndex)
                throw new SsTemplateException("Too many colors - not enough room in palette!");

            if (workbook instanceof HSSFWorkbook) {
                if (palette == null) palette = ((HSSFWorkbook)workbook).getCustomPalette();
                palette.setColorAtIndex(currentColorIndex, (byte) triplet[0], (byte) triplet[1], (byte) triplet[2]);
            }

            index = new Short(currentColorIndex);
            colorMap.put(color, index);
            currentColorIndex++;
        }

        return index.shortValue();
    }

    public void setColorIndex() {

    }

    public int getMaxRowIndex () {
        return maxRowIndex;
    }

    public int getMaxColumnIndex () {
        return maxColumnIndex;
    }

    public void setBackgroundColor (short[] triplet) {
        this.backgroundColor = triplet;
    }

    public short[] getBackgroundColor () {
        return backgroundColor;
    }

    public int nextPageBreak (int row) {
        if ((firstPageBreak <= 0) || (nextPageBreak <= 0)) return Short.MAX_VALUE;
        if (row < firstPageBreak) return firstPageBreak;
        return firstPageBreak + ((row - firstPageBreak) / nextPageBreak + 1) * nextPageBreak;
    }

    private class Color {
        private short[] triplet;

        public Color(short[] triplet) {
            this.triplet = triplet;
        }

        public int hashCode () {
            return triplet[0] * 256 * 256 + triplet[1] * 256 + triplet[2];
        }

        public boolean equals (Object obj) {
            return Arrays.equals(triplet, ((Color) obj).triplet);
        }
    }

    private class FontKey {
        String name = null;
        short  fontHeight, color, typeOffset;
        boolean italic, strikeout, bold;
        byte    underline;

        public FontKey(String name, short fontHeight, short color, boolean bold,
                boolean italic, boolean strikeout, byte underline, short typeOffset) {
            this.name = name;
            this.fontHeight = fontHeight;
            this.color = color;
            this.bold = bold;
            this.italic = italic;
            this.strikeout = strikeout;
            this.underline = underline;
            this.typeOffset = typeOffset;
        }

        public boolean equals (Object other) {
            if ((other == null) || (!(other instanceof FontKey))) return false;
            FontKey font = (FontKey) other;
            return (new EqualsBuilder()).append(name, font.name)
                    .append(fontHeight, font.fontHeight).append(color, font.color).append(
                            bold, font.bold).append(italic, font.italic).append(
                            strikeout, font.strikeout).append(underline, font.underline).append(
                            typeOffset, font.typeOffset).isEquals();
        }

        public int hashCode () {
            return (new HashCodeBuilder()).append(name).append(fontHeight).append(color).append(
                    bold).append(italic).append(strikeout).append(underline).append(
                    typeOffset).toHashCode();
        }

        public void setFontProperties (Font font) {
            font.setFontName(name);
            font.setFontHeight(fontHeight);
            font.setColor(color);
            font.setBold(bold);
            font.setItalic(italic);
            font.setStrikeout(strikeout);
            font.setUnderline(underline);
            font.setTypeOffset(typeOffset);
        }
    }

    private class CachedStyle {
        public StyleData styleData;
        public CellStyle style;

        public CachedStyle(StyleData styleData, CellStyle style) {
            this.styleData = styleData;
            this.style = style;
        }
    }

    public static class RemoveOnGetMap implements Map {
        private Map map;

        public RemoveOnGetMap(Map map) {
            this.map = map;
        }

        public int size () {
            return map.size();
        }

        public void clear () {
            map.clear();
        }

        public boolean isEmpty () {
            return map.isEmpty();
        }

        public boolean containsKey (Object key) {
            return map.containsKey(key);
        }

        public boolean containsValue (Object value) {
            return map.containsValue(value);
        }

        public Collection values () {
            return map.values();
        }

        public void putAll (Map t) {
            map.putAll(t);
        }

        public Set entrySet () {
            return map.entrySet();
        }

        public Set keySet () {
            return map.keySet();
        }

        public Object get (Object key) {
            return map.remove(key);
        }

        public Object remove (Object key) {
            return map.remove(key);
        }

        public Object put (Object key, Object value) {
            return map.put(key, value);
        }
    }

}
