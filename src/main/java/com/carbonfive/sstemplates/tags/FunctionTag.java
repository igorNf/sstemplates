package com.carbonfive.sstemplates.tags;

import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;
import java.util.StringTokenizer;

import com.carbonfive.sstemplates.SsTemplateContext;
import com.carbonfive.sstemplates.SsTemplateException;


/**
 * 
 * @author sivoh
 * @version $REVISION
 */
@SuppressWarnings("unchecked")
public class FunctionTag extends BaseTag {
    
    private String name           = null;
    private String clazz          = null;
    private String methodName     = null;
    private String parameterTypes = null;

    public void render (SsTemplateContext context) throws SsTemplateException {
        if (name == null) throw new SsTemplateException("Function tag must have a name attribute");
        String parsedName = (String) parseExpression(name, String.class, context);

        if (clazz == null)
            throw new SsTemplateException("Function tag must have a class attribute");
        String parsedClazz = (String) parseExpression(clazz, String.class, context);

        if (methodName == null)
            throw new SsTemplateException("Function tag must have a methodName attribute");
        String parsedMethodName = (String) parseExpression(methodName, String.class, context);

        String parsedParameterTypes = null;
        if (parameterTypes != null)
            parsedParameterTypes = (String) parseExpression(parameterTypes, String.class, context);

        try {
            Class c = classForName(parsedClazz);
            List parameters = new ArrayList();
            if (parsedParameterTypes != null) {
                for (StringTokenizer tok = new StringTokenizer(parsedParameterTypes, " "); tok
                        .hasMoreTokens();) {
                    Class p = classForName(tok.nextToken());
                    parameters.add(p);
                }
            }

            Method m = c.getMethod(parsedMethodName, (Class[]) parameters.toArray(new Class[] {}));
            context.registerMethod(parsedName, m);
        } catch (Exception e) {
            throw new SsTemplateException(e);
        }

        renderChildren(context);
    }

    private Class classForName (String clazz) throws ClassNotFoundException {
        if (clazz.equals("boolean")) return boolean.class;
        if (clazz.equals("byte")) return byte.class;
        if (clazz.equals("char")) return char.class;
        if (clazz.equals("double")) return double.class;
        if (clazz.equals("float")) return float.class;
        if (clazz.equals("int")) return int.class;
        if (clazz.equals("long")) return long.class;
        if (clazz.equals("short")) return short.class;
        return Class.forName(clazz);
    }

    public String getName () {
        return name;
    }

    public void setName (String name) {
        this.name = name;
    }

    public String getClassName () {
        return clazz;
    }

    public void setClassName (String clazz) {
        this.clazz = clazz;
    }

    public String getMethodName () {
        return methodName;
    }

    public void setMethodName (String methodName) {
        this.methodName = methodName;
    }

    public String getParameterTypes () {
        return parameterTypes;
    }

    public void setParameterTypes (String parameterTypes) {
        this.parameterTypes = parameterTypes;
    }
}
