package com.carbonfive.sstemplates.tags;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.Iterator;
import java.util.Map;

import com.carbonfive.sstemplates.SsTemplateContext;
import com.carbonfive.sstemplates.SsTemplateException;


/**
 * 
 * @author sivoh
 * @version $REVISION
 */
@SuppressWarnings("unchecked")
public class ForEachTag extends BaseTag {

    protected String var;
    protected String varStatus;
    protected String items;
    protected String begin = "0";
    protected String end;
    protected String step  = "1";

	public void render (SsTemplateContext context) throws SsTemplateException {
        // if ( var == null )
        // throw new
        // SsTemplateException("ForEach tag must have a 'var' attribute");

        // if (items == null)
        // throw new
        // SsTemplateException("ForEach tag must have an 'items' attribute");

        String parsedVar = (String) parseExpression(var, String.class, context);
        String parsedVarStatus = (String) parseExpression(varStatus, String.class, context);
        Integer parsedBeginValue = (Integer) parseExpression(begin, Integer.class, context);
        Integer parsedEndValue = (Integer) parseExpression(end, Integer.class, context);
        Integer parsedStepValue = (Integer) parseExpression(step, Integer.class, context);

        int parsedBegin = parsedBeginValue == null ? 0 : parsedBeginValue.intValue();
        int parsedEnd = parsedEndValue == null ? 0 : parsedEndValue.intValue();
        int parsedStep = parsedStepValue == null ? 0 : parsedStepValue.intValue();

        Object oldVar = context.getPageVariable(parsedVar);
        Object oldVarStatus = context.getPageVariable(parsedVarStatus);

        Collection c = (items == null ? new InfiniteList() : findCollection(context));
        if (end == null) parsedEnd = c.size();
        VarStatus varStatusObj = new VarStatus(c.size(), parsedStep, parsedBegin, parsedEnd);
        int i = 1;
        Iterator it = c.iterator();

        while (i <= parsedBegin && it.hasNext()) {
            it.next();
            i++;
        }
        while (i <= parsedEnd && it.hasNext()) {
            Object current = it.next();
            if (var != null) context.setPageVariable(parsedVar, current);
            if (varStatus != null) context.setPageVariable(parsedVarStatus, varStatusObj);
            renderChildren(context);

            for (int s = 1; s < parsedStep && it.hasNext(); s++) {
                it.next();
                i++;
            }
            varStatusObj.inc();
        }
        context.unsetPageVariable(parsedVar, oldVar);
        context.unsetPageVariable(parsedVarStatus, oldVarStatus);
    }

    @SuppressWarnings("serial")
	private static class InfiniteList extends ArrayList {
        public Iterator iterator () {
            return new InfiniteIterator();
        }
    }

    private static class InfiniteIterator implements Iterator {
        public boolean hasNext () {
            return true;
        }

        public void remove () {
            return;
        }

        public Object next () {
            return null;
        }
    }

    private Collection findCollection (SsTemplateContext context) throws SsTemplateException {

        Object obj = parseExpression(items, Object.class, context);
        if (obj == null)
            throw new SsTemplateException("Cannot find '" + items + "' in the page context");

        if (obj instanceof Collection) return (Collection) obj;
        else if (obj instanceof Map) {
            return ((Map) obj).entrySet();
        } else if (obj.getClass().isArray()) return Arrays.asList((Object[]) obj);
        else throw new SsTemplateException("'" + items
                + "' does not resolve to a collection, map or array");
    }

    public void setVar (String var) {
        this.var = var;
    }

    public void setItems (String items) {
        this.items = items;
    }

    public void setVarStatus (String varStatus) {
        this.varStatus = varStatus;
    }

    public void setBegin (String begin) {
        this.begin = begin;
    }

    public void setEnd (String end) {
        this.end = end;
    }

    public void setStep (String step) {
        this.step = step;
    }

    public static class VarStatus {
        private int size;
        private int step;
        private int begin;
        private int end;
        private int index;
        private int count;

        public VarStatus(int size, int step, int begin, int end) {
            this.size = size;
            this.step = step;
            this.begin = begin;
            this.end = end;
            this.index = begin;
            this.count = 1;
        }

        protected void inc () {
            this.count++;
            this.index += step;
        }

        public int getIndex () {
            return index;
        }

        public int getCount () {
            return count;
        }

        public boolean isFirst () {
            return index == begin;
        }

        public boolean isLast () {
            return index + step >= end || index + step >= size;
        }
    }
}
