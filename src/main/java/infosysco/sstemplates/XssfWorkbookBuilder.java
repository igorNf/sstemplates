package infosysco.sstemplates;


import java.util.Map;

import org.apache.poi.ss.usermodel.Workbook;

import com.carbonfive.sstemplates.SsTemplateContext;
import com.carbonfive.sstemplates.SsTemplateContextImpl;
import com.carbonfive.sstemplates.SsTemplateException;
import com.carbonfive.sstemplates.tags.WorkbookTag;

public class XssfWorkbookBuilder {

    private final WorkbookTag renderTree;

    XssfWorkbookBuilder(final WorkbookTag tree) {
        renderTree = tree;
    }

    public Workbook build(Map<String, Object> context) throws SsTemplateException {

        SsTemplateContext templateContext = new SsTemplateContextImpl(context);
        renderTree.renderNew(templateContext);
        return templateContext.getWorkbook();
    }
}
