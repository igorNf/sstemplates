package com.carbonfive.sstemplates;

import java.lang.reflect.Method;

import javax.servlet.jsp.el.FunctionMapper;
import javax.servlet.jsp.el.VariableResolver;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import org.apache.poi.ss.util.CellRangeAddress;

import com.carbonfive.sstemplates.hssf.CellAccumulator;
import com.carbonfive.sstemplates.hssf.StyleData;


@SuppressWarnings("deprecation")
public interface SsTemplateContext extends VariableResolver, FunctionMapper {

    Object setPageVariable (String key, Object value);

    void unsetPageVariable (String key, Object oldValue);

    Object getPageVariable (String key);

    Object resolveVariable (String name);

    Font createFont (String name, short fontHeight, short color, boolean bold,
            boolean italic, boolean strikeout, byte underline, short typeOffset);

    void incrementCellIndex ();

    void incrementRowIndex ();

    CellRangeAddress getRegionForCurrentLocation ();

    String addStyleData (String name, StyleData data);

    CellStyle getNamedStyle (String name) throws SsTemplateException;

    StyleData getNamedStyleData (String name) throws SsTemplateException;

    boolean hasCachedStyleData (String name);

    Workbook getWorkbook ();

    void setWorkbook (Workbook workbook);

    Sheet getSheet ();

    void setSheet (Sheet sheet);

    Row getRow ();

    void setRow (Row row);

    int getRowIndex ();

    void setRowIndex (int rowIndex);

    int getColumnIndex ();

    void setColumnIndex (int columnIndex);

    String getCurrentStyle ();

    void setCurrentStyle (String currentStyle);

    CellAccumulator getNamedAccumulator (String name);

    void registerMethod (String name, Method m);

    // no prefix support
    Method resolveFunction (String prefix, String name);

    Object getCustomValue (Object key);

    void setCustomValue (Object key, Object value);

    short getColorIndex (short[] triplet) throws SsTemplateException;

    void setBackgroundColor (short[] triplet);

    short[] getBackgroundColor ();

    public int getMaxRowIndex ();

    public int getMaxColumnIndex ();

    public void setPageBreaks (int firstPageBreak, int nextPageBreak);

    public int nextPageBreak (int row);
}
