package services;

import infosysco.sstemplates.HssfWorkbookBuilder;
import infosysco.sstemplates.ReportMapFactoryBean;
import infosysco.sstemplates.XssfWorkbookBuilder;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;
import java.util.HashMap;
import java.util.Map;


public class Main {
    public static void main(String[] args) throws IOException {
        //SsTemplatesService ssTemplatesService = new SsTemplatesService();
        Map<String, Object> map = new HashMap<>();

        map.put("currency", "RUB");
        map.put("employee", "RUB");
        // byte[] bytes = ssTemplatesService.makeReport("test_template.sst.xml", map, false);
        FileOutputStream fos = null;
        try {
            byte[] bytes = new Main().makeReport("test_template.sst.xml", map, false);

            fos = new FileOutputStream(new File("/home/nefedov/Documents/test.xlsx"));

            fos.write(bytes);
        } catch (Exception ex) {
            ex.printStackTrace();
        } finally {
            fos.close();
        }
    }

    byte[] makeReport(String reportName, Map<String, Object> context, boolean toXlsx) throws Exception {

        InputStream is = getClass().getClassLoader().getResourceAsStream(reportName);
        // XssfWorkbookBuilder builder = ReportMapFactoryBean.makeXssfWorkbookBuilder(is);
        HssfWorkbookBuilder builder = ReportMapFactoryBean.makeHssfWorkbookBuilder(is);

        Workbook workBook = builder.build(context);
        return getWorkbookBytes(workBook, toXlsx);
    }

    byte[] getWorkbookBytes(Workbook workBook, boolean toXlsx) throws Exception{

        final ByteArrayOutputStream baos = new ByteArrayOutputStream(4096);

        workBook.write(baos);

        return baos.toByteArray();
    }
}
