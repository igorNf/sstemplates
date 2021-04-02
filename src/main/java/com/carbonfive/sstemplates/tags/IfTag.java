package com.carbonfive.sstemplates.tags;

import com.carbonfive.sstemplates.SsTemplateContext;
import com.carbonfive.sstemplates.SsTemplateException;


/**
 * 
 * @author sivoh
 * @version $REVISION
 */
public class IfTag extends BaseTag {

    protected String test;
    private String   var;
    private String   scope = "page";

    public void render (SsTemplateContext context) throws SsTemplateException {
        if (test == null) throw new SsTemplateException("If tag must have a 'test' attribute");

        if (!"page".equals(scope))
            throw new SsTemplateException("Set tag 'scope' attribute can only be 'page'");

        Boolean parsedTest = (Boolean) parseExpression(test, Boolean.class, context);

        if (var != null) {
            String parsedVar = (String) parseExpression(var, String.class, context);
            context.setPageVariable(parsedVar, parsedTest);
            renderChildren(context);
        } else {
            if (parsedTest.booleanValue()) renderChildren(context);
        }
    }

    public void setTest (String test) {
        this.test = test;
    }

    public void setVar (String var) {
        this.var = var;
    }

    public void setScope (String scope) {
        this.scope = scope;
    }
}
