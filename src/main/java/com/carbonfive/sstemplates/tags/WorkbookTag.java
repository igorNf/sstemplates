package com.carbonfive.sstemplates.tags;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.carbonfive.sstemplates.SsTemplateContext;
import com.carbonfive.sstemplates.SsTemplateException;


public class WorkbookTag extends BaseTag {

    private String              bgcolor = null;

    public void render (SsTemplateContext context) throws SsTemplateException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        context.setWorkbook(workbook);

        if (bgcolor != null) {
            short triplet[];
            String parsedColor = (String) parseExpression(bgcolor, String.class, context);
            if (parsedColor.startsWith("#")) triplet = StyleTag.parseColor(parsedColor);
            else throw new SsTemplateException("Can't parse background color '" + parsedColor + "'");

            context.setBackgroundColor(triplet);
        }

        renderChildren(context);
    }

    public void renderNew (SsTemplateContext context) throws SsTemplateException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        context.setWorkbook(workbook);

        renderChildren(context);
    }

    public String getBgcolor () {
        return bgcolor;
    }

    public void setBgcolor (String bgcolor) {
        this.bgcolor = bgcolor;
    }
}
