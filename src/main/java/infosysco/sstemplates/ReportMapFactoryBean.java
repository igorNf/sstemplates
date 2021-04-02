package infosysco.sstemplates;

import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.commons.digester3.Digester;
import org.springframework.beans.factory.FactoryBean;
import org.springframework.beans.factory.InitializingBean;
import org.springframework.core.io.Resource;
import org.xml.sax.SAXException;

import com.carbonfive.sstemplates.tags.CellTag;
import com.carbonfive.sstemplates.tags.DefaultStyleTag;
import com.carbonfive.sstemplates.tags.ForEachTag;
import com.carbonfive.sstemplates.tags.FunctionTag;
import com.carbonfive.sstemplates.tags.IfTag;
import com.carbonfive.sstemplates.tags.RowBreakTag;
import com.carbonfive.sstemplates.tags.RowTag;
import com.carbonfive.sstemplates.tags.SetTag;
import com.carbonfive.sstemplates.tags.SheetTag;
import com.carbonfive.sstemplates.tags.SsTemplateTag;
import com.carbonfive.sstemplates.tags.StyleTag;
import com.carbonfive.sstemplates.tags.WorkbookTag;

import org.springframework.core.io.DefaultResourceLoader;
import org.springframework.core.io.Resource;
import org.springframework.core.io.ResourceLoader;


public class ReportMapFactoryBean implements FactoryBean, InitializingBean {

    @SuppressWarnings("unchecked")
    private static Class<SsTemplateTag>[] supportedTagClasses = new Class[] { CellTag.class,
            DefaultStyleTag.class, ForEachTag.class, FunctionTag.class, IfTag.class,
            RowBreakTag.class, RowTag.class, SetTag.class, SheetTag.class, StyleTag.class };

    private Map<String, HssfWorkbookBuilder> reportMap;

    private Map<String, Resource>            resourceMap;

    private String resourceBasePath = "";

    private ResourceLoader resourceLoader = new DefaultResourceLoader();

    @Override
    public Object getObject() throws Exception {
        return reportMap;
    }

    @Override
    public Class<?> getObjectType () {
        return Map.class;
    }

    @Override
    public final boolean isSingleton () {
        return true;
    }

    @Override
    public void afterPropertiesSet () throws Exception {
        this.reportMap = createInstance();
    }

    public void setResourceMap (Map<String, String> resourceMap) {
        Map<String, Resource> _resourceMap = new HashMap();
        for (Map.Entry<String, String> resource: resourceMap.entrySet()) {
            _resourceMap.put(resource.getKey(), getResource(resource.getValue()));
        }

        this.resourceMap = _resourceMap;
    }

    protected Resource getResource(String location) {
        return this.resourceLoader.getResource(this.resourceBasePath + location);
    }

    public void setResourceBasePath(String resourceBasePath) {
        this.resourceBasePath = (resourceBasePath != null ? resourceBasePath : "");
    }

    // **************************************************************************

    private Map<String, HssfWorkbookBuilder> createInstance () throws IOException, SAXException {
        Map<String, HssfWorkbookBuilder> result = new HashMap<String, HssfWorkbookBuilder>();

        for (Entry<String, Resource> entry : resourceMap.entrySet()) {
            result.put(entry.getKey(), makeHssfWorkbookBuilder(entry.getValue().getInputStream()));
        }
        return result;
    }

    public static HssfWorkbookBuilder makeHssfWorkbookBuilder (InputStream is) throws IOException,
            SAXException {

        Digester digester = _makeDigester(is);
        return new HssfWorkbookBuilder((WorkbookTag) digester.getRoot());
    }

    public static XssfWorkbookBuilder makeXssfWorkbookBuilder (InputStream is) throws IOException, SAXException {

        Digester digester = _makeDigester(is);
        return new XssfWorkbookBuilder((WorkbookTag) digester.getRoot());
    }

    private static Digester _makeDigester (InputStream is) throws IOException, SAXException {

        Digester digester = new Digester();
        digester.setClassLoader(ReportMapFactoryBean.class.getClassLoader());
        digester.setValidating(false);
        digester.addSetProperties("workbook");

        for (Class<SsTemplateTag> clazz : supportedTagClasses) {

            String pattern = "*/" + _tagNameForClass(clazz);

            digester.addObjectCreate(pattern, clazz);
            digester.addSetProperties(pattern);
            digester.addSetNext(pattern, "addTag");
        }

        // add special rule to add content to cell
        digester.addCallMethod("*/cell", "setContents", 0);
        digester.push(new WorkbookTag());
        try {
            digester.parse(is);
        } finally {
            is.close();
        }

        return digester;
    }

    private static String _tagNameForClass (Class<?> clazz) {

        String tagClassName = clazz.getSimpleName().replaceAll("Tag", "");
        return tagClassName.substring(0, 1).toLowerCase().concat(tagClassName.substring(1));
    }
}
