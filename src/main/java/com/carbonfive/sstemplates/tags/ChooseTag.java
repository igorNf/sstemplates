package com.carbonfive.sstemplates.tags;

import com.carbonfive.sstemplates.SsTemplateContext;
import com.carbonfive.sstemplates.SsTemplateException;


public class ChooseTag extends BaseTag {

    public void render (SsTemplateContext context) throws SsTemplateException {

        renderChildren(context);
    }
}
