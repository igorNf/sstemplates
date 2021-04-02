package com.carbonfive.sstemplates.tags;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.logging.Logger;

import javax.servlet.jsp.el.ELException;

import org.apache.commons.el.ExpressionEvaluatorImpl;

import com.carbonfive.sstemplates.SsTemplateContext;
import com.carbonfive.sstemplates.SsTemplateException;


/**
 *
 * @author sivoh
 * @version $REVISION
 */
public abstract class BaseTag implements SsTemplateTag {

    protected static final Logger log       = Logger.getLogger(BaseTag.class.getName());

    List<SsTemplateTag>           childTags = new ArrayList<SsTemplateTag>();
    ExpressionEvaluatorImpl       evaluator = new ExpressionEvaluatorImpl();

    protected void renderChildren (SsTemplateContext context) throws SsTemplateException {
        renderChildren(context, childTags);
    }

    protected void renderChildren (SsTemplateContext context, Collection<SsTemplateTag> children)
            throws SsTemplateException {

        for (SsTemplateTag tag : children) {
            tag.render(context);
        }
    }

    public void addTag (SsTemplateTag tag) {
        childTags.add(tag);
    }

    public List<SsTemplateTag> getChildTags () {
        return childTags;
    }

    public Object parseExpression (String expression, Class<?> expectedType,
            SsTemplateContext context) throws SsTemplateException {
        if (expression == null) return null;

        try {
            log.fine("Evaluating expression " + expression);
            return evaluator.evaluate(expression, expectedType, context, context);
        } catch (ELException ele) {
            throw new SsTemplateException("Error parsing expression " + expression, ele);
        }
    }

    public Integer parseInteger (String expression, SsTemplateContext context)
            throws SsTemplateException {
        return (Integer) parseExpression(expression, Integer.class, context);
    }

    public int parseInt (String expression, SsTemplateContext context) throws SsTemplateException {
        return parseInteger(expression, context).intValue();
    }
}
