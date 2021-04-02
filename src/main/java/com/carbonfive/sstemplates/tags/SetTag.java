package com.carbonfive.sstemplates.tags;

import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;

import com.carbonfive.sstemplates.SsTemplateContext;
import com.carbonfive.sstemplates.SsTemplateException;


public class SetTag extends BaseTag {

    protected String var;
    protected String value;
    protected String target;
    protected String property;
    protected String scope = "page";

    public void setVar (String var) {

        this.var = var;
    }

    public void setValue (String value) {

        this.value = value;
    }

    public void setTarget (String target) {

        this.target = target;
    }

    public void setProperty (String property) {

        this.property = property;
    }

    public void setScope (String scope) {

        this.scope = scope;
    }

    public void render (SsTemplateContext context) throws SsTemplateException {

        if (var == null && target == null)
            throw new SsTemplateException("Set tag must have a 'var' or 'target' attribute");

        if (var != null && target != null)
            throw new SsTemplateException("Set tag cannot have both 'var' and 'target' attributes");

        if (value == null) throw new SsTemplateException("Set tag must have a 'value' attribute");

        if (!"page".equals(scope))
            throw new SsTemplateException("Set tag 'scope' attribute can only be 'page'");

        Object parsedValue = (Object) parseExpression(value, Object.class, context);

        if (var != null) {
            String parsedVar = (String) parseExpression(var, String.class, context);
            if (context.getPageVariable(parsedVar) != null)
                throw new SsTemplateException("Page variable already exists: " + parsedVar);

            context.setPageVariable(parsedVar, parsedValue);
        } else if (target != null) {
            String parsedTarget = (String) parseExpression(target, String.class, context);
            if (context.getPageVariable(parsedTarget) == null)
                throw new SsTemplateException("Cannot find page variable: " + parsedTarget);

            if (property == null) context.setPageVariable(parsedTarget, parsedValue);
            else {
                String parsedProperty = (String) parseExpression(property, String.class, context);
                Object targetObject = context.getPageVariable(parsedTarget);
                try {
                    PropertyDescriptor pd = new PropertyDescriptor(parsedProperty, targetObject
                            .getClass());
                    pd.getWriteMethod().invoke(targetObject, new Object[] { parsedValue });
                } catch (IntrospectionException ie) {
                    throw new SsTemplateException("Page variable '" + parsedTarget
                            + "' has no property '" + parsedProperty + "'");
                } catch (Exception e) {
                    throw new SsTemplateException("Error setting '" + parsedTarget + "."
                            + parsedProperty + "'", e);
                }
            }
        }
        renderChildren(context);
    }
}
