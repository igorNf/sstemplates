package infosysco.sstemplates;

import java.util.Map;

import org.apache.poi.ss.usermodel.Workbook;

import com.carbonfive.sstemplates.SsTemplateContext;
import com.carbonfive.sstemplates.SsTemplateContextImpl;
import com.carbonfive.sstemplates.SsTemplateException;
import com.carbonfive.sstemplates.tags.WorkbookTag;


public class HssfWorkbookBuilder {

    private final WorkbookTag renderTree;

    HssfWorkbookBuilder(final WorkbookTag tree) {
        renderTree = tree;
    }

    public Workbook build (Map<String, Object> context) throws SsTemplateException {

        SsTemplateContext templateContext = new SsTemplateContextImpl(context);
        renderTree.render(templateContext);
        return templateContext.getWorkbook();
    }
}
