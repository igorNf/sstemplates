package com.carbonfive.sstemplates.tags;

import org.apache.poi.ss.usermodel.Sheet;

import com.carbonfive.sstemplates.SsTemplateContext;
import com.carbonfive.sstemplates.SsTemplateException;
import com.carbonfive.sstemplates.hssf.StyleData;


/**
 *
 * @author sivoh
 * @version $REVISION
 */
public class DefaultStyleTag extends StyleTag {

    protected void renderChildren (SsTemplateContext context) throws SsTemplateException {
        StyleData styleData = context.getNamedStyleData(context.getCurrentStyle());

        if (context.getSheet() != null)
            throw new SsTemplateException("Must define defaultstyle before creating sheets");

        Sheet sheet = context.getWorkbook().createSheet();
        styleData.setStyleAttributes(sheet.createRow(0).createCell(0).getCellStyle(), context);
        context.getWorkbook().removeSheetAt(0);

        super.renderChildren(context);
    }
}
