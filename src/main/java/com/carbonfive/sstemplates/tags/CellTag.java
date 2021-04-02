package com.carbonfive.sstemplates.tags;

import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.StringTokenizer;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;

import com.carbonfive.sstemplates.SsTemplateContext;
import com.carbonfive.sstemplates.SsTemplateException;
import com.carbonfive.sstemplates.hssf.CellAccumulator;
import com.carbonfive.sstemplates.hssf.StyleData;

/**
 *
 * @author sivoh
 * @version $REVISION
 */
@SuppressWarnings({"unchecked", "deprecation"})
public class CellTag extends BaseTag {

    private String column         = null;
    private String relativeColumn = null;
    private String contents       = null;
    private String type           = null;
    private String parsedType     = null;
    private String colspan        = null;
    private String rowspan        = null;
    private String style          = null;
    private String accumulator    = null;
    private String paginate       = null;

    public void render (SsTemplateContext context) throws SsTemplateException {
        if (context.getRow() == null)
            throw new SsTemplateException("Cell tag must be within a row tag");

        int columnIndex = context.getColumnIndex();
        if (column != null) {
            columnIndex = ((Integer) parseExpression(column, Integer.class, context)).intValue();
            context.setColumnIndex(columnIndex);
        } else if (relativeColumn != null) {
            columnIndex = columnIndex
                    + ((Integer) parseExpression(relativeColumn, Integer.class, context))
                            .intValue();
            context.setColumnIndex(columnIndex);
        }

        // need to reset row because paginated cells might have moved it around
        Row row = context.getRow();
        int rowIndex = context.getRowIndex();
        createCell(context);
        context.setRow(row);
        context.setRowIndex(rowIndex);

        // accumulator get set here because we don't want pagination to cause
        // duplication
        if (accumulator != null) {
            Cell cell = context.getRow().getCell((short) columnIndex);
            setCellAccumulator(cell, context, accumulator);
        }

        context.incrementCellIndex();
    }

	private void createCell (SsTemplateContext context) throws SsTemplateException {
        CellRangeAddress region = createRegion(context);

        if ((paginate == null)
                || !((Boolean) parseExpression(paginate, Boolean.class, context)).booleanValue()) {
            createCell(context, region, true, true);
            return;
        }

        List regions = splitRegionForPagination(context, region);
        for (Iterator i = regions.iterator(); i.hasNext();) {
            CellRangeAddress sub = (CellRangeAddress) i.next();
            createCell(context, sub, sub.getFirstRow() == region.getFirstRow(),
                    sub.getLastRow() == region.getLastRow());
        }
    }

    private List splitRegionForPagination (SsTemplateContext context, CellRangeAddress region) {
        List regions = new ArrayList();
        int rowFrom = region.getFirstRow();
        while (region.getLastRow() >= context.nextPageBreak(rowFrom)) {
            regions.add(new CellRangeAddress(rowFrom, context.nextPageBreak(rowFrom) - 1, region.getFirstColumn(),
                     region.getLastColumn()));
            rowFrom = context.nextPageBreak(rowFrom);
        }
        regions.add(new CellRangeAddress(rowFrom, region.getLastRow(), region.getFirstColumn(), region
                .getLastColumn()));
        return regions;
    }

    private void createCell (SsTemplateContext context, CellRangeAddress region, boolean showTopBorder,
            boolean showBottomBorder) throws SsTemplateException {
        int rowIndex = region.getFirstRow();
        int columnIndex = region.getFirstColumn();

        context.setRowIndex(rowIndex);
        Row row = context.getSheet().getRow(rowIndex);
        if (row == null) row = context.getSheet().createRow(rowIndex);
        context.setRow(row);

        Cell cell = context.getRow().getCell((short) columnIndex);
        if (cell == null) cell = context.getRow().createCell((short) columnIndex);

        // [20050914 jah] fix for cyrillic
        // ??? cell.setEncoding(HSSFCell.ENCODING_UTF_16);

        cell.setCellType(findCellType(context));

        setCellContents(cell, context);

        setCellStyle(cell, context, showTopBorder, showBottomBorder);

        if ((colspan != null) || (rowspan != null)) {
            context.getSheet().addMergedRegion(region);
            createRegionBorders(region, cell, context);
        }
    }

    private void setCellAccumulator (Cell cell, SsTemplateContext context, String accumulator)
            throws SsTemplateException {
        for (StringTokenizer tok = new StringTokenizer(accumulator, " "); tok.hasMoreTokens();) {
            CellAccumulator acc = context.getNamedAccumulator((String) parseExpression(tok
                    .nextToken(), String.class, context));
            acc.addCell(cell, context.getRowIndex(), context.getColumnIndex());
        }
    }

    private void setCellStyle (Cell cell, SsTemplateContext context, boolean showTopBorder,
            boolean showBottomBorder) throws SsTemplateException {
        String styleName = context.getCurrentStyle();
        if (style != null)
            styleName = appendStyles(styleName, (String) parseExpression(style, String.class,
                    context));
        if (!showTopBorder) styleName = appendStyles(styleName, "_noTopBorder");
        if (!showBottomBorder) styleName = appendStyles(styleName, "_noBottomBorder");

        if (styleName.length() > 0) {
            cell.setCellStyle(context.getNamedStyle(styleName));

            // set special parameters not contained in style
            StyleData data = context.getNamedStyleData(styleName);
            if ((data.getAutoColumnWidth()) && (cell.getCellType() == CellType.STRING)
                    && (cell.getStringCellValue() != null)) {
                int width = context.getSheet().getColumnWidth((short) context.getColumnIndex());
                context.getSheet().setColumnWidth(
                        (short) context.getColumnIndex(),
                        (short) Math.min(256 * 100, Math.max(width, 256 * (cell
                                .getStringCellValue().length() + 2))));
            } else if (data.getColumnWidth() != null)
                context.getSheet().setColumnWidth((short) context.getColumnIndex(),
                        data.getColumnWidth().shortValue());

            if (data.getRowHeight() != null)
                context.getRow().setHeight(data.getRowHeight().shortValue());
        }
    }

    private String appendStyles (String styleName, String parsedStyle) {
        if (styleName.length() > 0) styleName += " " + parsedStyle;
        else styleName = parsedStyle;
        return styleName;
    }

    private CellRangeAddress createRegion (SsTemplateContext context) throws SsTemplateException {
        short parsedColspan = 1;
        int parsedRowspan = 1;

        CellRangeAddress region = new CellRangeAddress(context.getRowIndex(), context
                .getRowIndex(), context.getColumnIndex(), context.getColumnIndex());

        if (colspan != null) {
            parsedColspan = ((Integer) parseExpression(colspan, Integer.class, context))
                    .shortValue();
            region.setLastColumn(context.getColumnIndex() + parsedColspan - 1);
        }

        if (rowspan != null) {
            parsedRowspan = parseInt(rowspan, context);
            region.setLastRow(context.getRowIndex() + parsedRowspan - 1);
        }

        return region;
    }

    private void createRegionBorders (CellRangeAddress region, Cell cell, SsTemplateContext context)
            throws SsTemplateException {
        int parsedColspan = region.getLastColumn() - region.getFirstColumn() + 1;
        int parsedRowspan = region.getLastRow() - region.getFirstRow() + 1;

        CellStyle style = cell.getCellStyle();
        if ((style.getBorderTop() != BorderStyle.NONE)
                || (style.getBorderBottom() != BorderStyle.NONE)
                || (style.getBorderRight() != BorderStyle.NONE)
                || (style.getBorderLeft() != BorderStyle.NONE)) {
            for (short i = 0; i < parsedColspan; i++) {
                for (int j = 0; j < parsedRowspan; j++) {
                    if ((i > 0) || (j > 0)) {
                        boolean leftBorder = (i == 0)
                                && (style.getBorderLeft() != BorderStyle.NONE);
                        boolean rightBorder = (i == parsedColspan - 1)
                                && (style.getBorderRight() != BorderStyle.NONE);
                        boolean topBorder = (j == 0)
                                && (style.getBorderTop() != BorderStyle.NONE);
                        boolean bottomBorder = (j == parsedRowspan - 1)
                                && (style.getBorderBottom() != BorderStyle.NONE);

                        if (leftBorder || rightBorder || topBorder || bottomBorder) {
                            String styleName = "!!!regionBorder-"
                                    + (leftBorder ? "L" + style.getBorderLeft() : "")
                                    + (rightBorder ? "R" + style.getBorderRight() : "")
                                    + (topBorder ? "T" + style.getBorderTop() : "")
                                    + (bottomBorder ? "B" + style.getBorderBottom() : "");
                            StyleData styleData = null;
                            if (!context.hasCachedStyleData(styleName)) {
                                styleData = new StyleData();
                                if (topBorder) {
                                    styleData.put("borderTop", new Integer(style.getBorderTop().getCode()));
                                    styleData.put("topBorderColor", new Integer(style.getTopBorderColor()));
                                }
                                if (bottomBorder) {
                                    styleData.put("borderBottom", new Integer(style.getBorderBottom().getCode()));
                                    styleData.put("bottomBorderColor", new Integer(style.getBottomBorderColor()));
                                }
                                if (rightBorder) {
                                    styleData.put("borderRight",
                                            new Integer(style.getBorderRight().getCode()));
                                    styleData.put("rightBorderColor", new Integer(style.getRightBorderColor()));
                                }
                                if (leftBorder) {
                                    styleData.put("borderLeft", new Integer(style.getBorderLeft().getCode()));
                                    styleData.put("leftBorderColor", new Integer(style.getLeftBorderColor()));
                                }

                                context.addStyleData(styleName, styleData);
                            }

                            Row row = context.getSheet().getRow(context.getRowIndex() + j);
                            if (row == null)
                                row = context.getSheet().createRow(context.getRowIndex() + j);
                            Cell newCell = row.createCell((short) (context.getColumnIndex() + i));

                            // [20050914 jah] fix for cyrillic
                            // ???
                            // newCell.setEncoding(HSSFCell.ENCODING_UTF_16);

                            newCell.setCellStyle(context.getNamedStyle(styleName));
                        }
                    }
                }
            }
        }
    }

    private void setCellContents (Cell cell, SsTemplateContext context)
            throws SsTemplateException {
        if ((contents != null) && (contents.length() != 0)) {
            if (cell.getCellType() == CellType.STRING) cell
                    .setCellValue((String) parseExpression(contents, String.class, context));
            else if (cell.getCellType() == CellType.NUMERIC) {
                if ((parsedType != null) && parsedType.equals("date")) cell
                        .setCellValue((Date) parseExpression(contents, Date.class, context));
                else cell.setCellValue(((Double) parseExpression(contents, Double.class, context))
                        .doubleValue());
            } else if (cell.getCellType() == CellType.BOOLEAN) cell
                    .setCellValue(((Boolean) parseExpression(contents, Boolean.class, context))
                            .booleanValue());
            else if (cell.getCellType() == CellType.FORMULA) cell
                    .setCellFormula((String) parseExpression(contents, String.class, context));
            else if (cell.getCellType() == CellType.BLANK) cell.setCellFormula(null);
        }
    }

    private CellType findCellType (SsTemplateContext context) throws SsTemplateException {
        CellType cellType = CellType.STRING;
        if ((type != null) && (type.length() > 0)) {
            parsedType = ((String) parseExpression(type, String.class, context)).toLowerCase();
            if ("blank".equals(parsedType)) cellType = CellType.BLANK;
            else if ("boolean".equals(parsedType)) cellType = CellType.BOOLEAN;
            else if ("error".equals(parsedType)) cellType = CellType.ERROR;
            else if ("formula".equals(parsedType)) cellType = CellType.FORMULA;
            else if ("numeric".equals(parsedType) || "date".equals(parsedType)) cellType = CellType.NUMERIC;
            else if ("string".equals(parsedType)) cellType = CellType.STRING;
            else throw new SsTemplateException("Invalid cell type: " + parsedType);
        } else if ((contents == null) || (contents.length() == 0)) {
            cellType = CellType.BLANK;
        }
        return cellType;
    }

    public String getColumn () {
        return column;
    }

    public void setColumn (String column) {
        this.column = column;
    }

    public String getRelativeColumn () {
        return relativeColumn;
    }

    public void setRelativeColumn (String relativeColumn) {
        this.relativeColumn = relativeColumn;
    }

    public String getContents () {
        return contents;
    }

    public void setContents (String contents) {
        this.contents = contents;
        if (this.contents != null) this.contents = this.contents.trim();
    }

    public String getType () {
        return type;
    }

    public void setType (String type) {
        this.type = type;
    }

    public String getColspan () {
        return colspan;
    }

    public void setColspan (String colspan) {
        this.colspan = colspan;
    }

    public String getRowspan () {
        return rowspan;
    }

    public void setRowspan (String rowspan) {
        this.rowspan = rowspan;
    }

    public String getStyle () {
        return style;
    }

    public void setStyle (String style) {
        this.style = style;
    }

    public String getAccumulator () {
        return accumulator;
    }

    public void setAccumulator (String accumulator) {
        this.accumulator = accumulator;
    }

    public String getPaginate () {
        return paginate;
    }

    public void setPaginate (String paginate) {
        this.paginate = paginate;
    }
}
