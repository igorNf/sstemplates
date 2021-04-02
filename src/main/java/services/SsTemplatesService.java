//package services;
//
//import infosysco.sstemplates.ReportMapFactoryBean
//import infosysco.sstemplates.XssfWorkbookBuilder;
//
//import java.io.ByteArrayOutputStream;
//import java.io.File;
//import java.io.FileInputStream;
//import java.io.InputStream;
//import java.util.Arrays;
//import java.util.List;
//import java.util.Map;
//
//class SsTemplatesService {
//
//    // static transactional = false;
//
//    def excelReportsMap
//
//    def builders
//
//    def makeExcelReport (reportName, Map<String, Object> context) {
//
//        def builder = excelReportsMap[reportName]
//
//        if (! builder) {
//            throw new RuntimeException("Формирование такого Excel отчета (${reportName}) не поддерживается")
//        }
//        builder.build(context);
//    }
//
//
//    byte[] getWorkbookBytes(workBook, toXlsx = false) {
//
//        final ByteArrayOutputStream baos = new ByteArrayOutputStream(4096);
//        // if (toXlsx){
//        //     def  newBook = new ExcelConverter(workBook).transform()
//        //     newBook.write(baos)
//        // } else {
//        workBook.write(baos)
//        //}
//
//        return baos.toByteArray()
//    }
//
//    byte[] makeReport(String reportName, Map<String, Object> context, boolean toXlsx = true) {
//        def builder = getBookBuilder( reportName)
//        def workBook = builder.build(context)
//        return getWorkbookBytes(workBook, toXlsx)
//    }
//
//    private XssfWorkbookBuilder getBookBuilder(String reportName) {
//        if (builders?.containsKey(reportName) && !Holders.config.ibank?.resources?.from?.catalog) {
//            return builders[reportName]
//        }
//        def reportFile = loadResource(reportName);
//        final InputStream targetStream = new FileInputStream(reportFile);
//        // def builder = ReportMapFactoryBean.makeHssfWorkbookBuilder(targetStream)
//        def builder = ReportMapFactoryBean.makeXssfWorkbookBuilder(targetStream);
//        if (builders == null) {
//            builders = [:]
//        }
//        builders[reportName] = builder
//        return builders[reportName]
//    }
//
//    private File loadResource(String templatePath){
//        File resourcesFile = null;
//        // Holders.config.ibank?.resources?.from?.catalog
//        if (false) {
//            File workFolder = new File("..//");
//            List<File> listFiles = Arrays.asList(workFolder.listFiles()); // workFolder.listFiles().toList();
//            listFiles.add( new File("..//../grails-inline-plugins/grails-ibank-common"));
//            for (File l : listFiles) {
//                String filePath = l.getAbsoluteFile().toString() + "/grails-app/conf/resources/$templatePath";
//                File resFile = new File(filePath);
//                if (resFile.exists() && resFile.isFile()) {
//                    resourcesFile = resFile;
//                    break;
//                }
//            }
//
//        } else {
//            String templateFileName = "classpath:/resources/${templatePath}";
//            resourcesFile = Holders.applicationContext.getResource(templateFileName).getFile();
//        }
//
//        return resourcesFile;
//    }
//
//}
