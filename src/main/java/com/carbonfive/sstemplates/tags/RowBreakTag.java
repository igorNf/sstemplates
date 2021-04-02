package com.carbonfive.sstemplates.tags;

import com.carbonfive.sstemplates.SsTemplateContext;
import com.carbonfive.sstemplates.SsTemplateException;


public class RowBreakTag extends BaseTag {

    private String condition = null;

    public void render (SsTemplateContext context) throws SsTemplateException {
        context.getSheet().setRowBreak(context.getRowIndex() - 1);
    }

    public String getCondition () {
        return condition;
    }

    public void setCondition (String condition) {
        this.condition = condition;
    }
}
