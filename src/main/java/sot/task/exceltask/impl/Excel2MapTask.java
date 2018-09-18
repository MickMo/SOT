package sot.task.exceltask.impl;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.dom4j.Document;
import org.dom4j.Node;
import sot.bean.BeanInfo;
import sot.bean.ProcessResult;
import sot.bean.exception.ExcelProcessingException;
import sot.bean.sheetdata.impl.SheetData;
import sot.task.exceltask.Excel2DataBaseTask;
import sot.util.BeanUtils;
import sot.util.XmlReadUtil;
import sot.util.excelprocessutil.ArrayToBeanArrayUtil;
import sot.util.excelprocessutil.ExcelTo2DArrayUtil;
import sot.util.fileutil.ReadExcelMappingUtil;
import sot.wordutil.logger.ExcelUtilLogger;

import java.io.IOException;
import java.util.*;

/**
 * Excel process Task <br>
 * Read data from a Excel File and convert those data into Map
 *
 * @author Monan
 *         created on 2018/7/27 11:31
 */
public class Excel2MapTask extends Excel2DataBaseTask<Map<String, Object>> {

    /**
     * create a new Excel process Task
     *
     * @param document XML Document with Task Configuration
     * @param targets  A String array which storing package locations of bean or full reference of a bean
     * @param workBook Excel Sheet,Data Source
     * @throws IOException Any IO exception before processing the excel file.
     */
    public Excel2MapTask(Document document, String[] targets, Workbook workBook) throws IOException {
        super(document, targets, workBook);
    }

    /**
     * create a new Excel process Task
     *
     * @param taskName specific which task to read from configuration.
     * @param targets  A String array which storing package locations of bean or full reference of a bean
     * @param workBook Excel Sheet,Data Source
     * @throws IOException Any IO exception before processing the excel file.
     */
    public Excel2MapTask(String taskName, String[] targets, Workbook workBook) throws IOException {
        super(taskName, targets, workBook);
    }

    /**
     * create a new Excel process Task
     *
     * @param taskName  specific which task to read from configuration.
     * @param targets   A String array which storing package locations of bean or full reference of a bean
     * @param workBook  Excel workBook,work as Data Source.
     * @param sheetName specific which sheet name that to process,process all available sheet if given <code>Null</code>,which request task mapping in
     *                  configuration.
     * @throws IOException Any IO exception before processing the excel file.
     */
    public Excel2MapTask(String taskName, String[] targets, Workbook workBook, String sheetName) throws IOException {
        super(taskName, targets, workBook, sheetName);
    }

    /**
     * Process the Task
     *
     * @throws ExcelProcessingException when processing the task,once the error reach the limits set in configuration,<br>
     *                                  the program will throw ExcelProcessingException immediately.<br>
     *                                  Error log can be read from the exception using {@code getExcelUtilLogger()} method.;
     * @throws Exception                Any encounter Exception during processing XML file or document
     */
    @Override
    public ProcessResult<Map<String, Object>> execute() throws ExcelProcessingException, Exception {
        try {
            if (document == null) {
                document = XmlReadUtil.getMappingForExcelImportExportTask();
            }
            //acquire task confidence node from xml file
            Node taskNode = ReadExcelMappingUtil.getTaskNode(document, taskName);
            String taskComment = ReadExcelMappingUtil.getTaskComment(taskNode);
            //data storing map
            Map<String, SheetData<Map<String, Object>>> sheetDatas = new HashMap<>();
            Map<String, List<BeanInfo>> beanInfosFromXMLMap = new HashMap<>();
            Map<String, Integer> taskHeaderRowNums = new HashMap<>();
            ExcelUtilLogger excelUtilLogger = new ExcelUtilLogger(taskName, taskComment);
            Iterator<Sheet> sheetIterator = null;
            if (workBook2007 == null) {
                sheetIterator = workBook2003.sheetIterator();
            } else {
                sheetIterator = workBook2007.sheetIterator();

            }
            while (sheetIterator.hasNext()) {
                Sheet sheet = sheetIterator.next();
                sheetName = sheet.getSheetName();
                //get row number which data start from
                int taskHeaderRowNum = ReadExcelMappingUtil.getTaskHeaderRowNum(taskNode, sheetName);
                //resolving bean mapping from configuration
                //TODO 转map数据应该不需要classname配置,处理过程从这里开始应该用不同的方法
                List<BeanInfo> beanInfosFromXML = ReadExcelMappingUtil.getBeanInfosFromXML(taskNode, excelUtilLogger, sheetName);
                //excel data bean bean data
                ExcelTo2DArrayUtil.sheetToBeanInfoArray(sheet, taskHeaderRowNum, beanInfosFromXML, excelUtilLogger);
                List<List<Object>> lists = ArrayToBeanArrayUtil.data2DarrayTo2DBeanArray(sheet, taskHeaderRowNum, beanInfosFromXML, excelUtilLogger);
                //Convert Bean Data to Map data
                List<Map<String, Object>> sheetMap = new ArrayList<>();
                for (List<Object> list : lists) {
                    Map<String, Object> rowMap = new HashMap<>();
                    for (Object o : list) {
                        rowMap.putAll(BeanUtils.transBean2Map(o));
                    }
                    sheetMap.add(rowMap);
                }
                //saving process result
                SheetData<Map<String, Object>> sheetData = new SheetData<Map<String, Object>>(sheet, sheetMap);
                sheetDatas.put(sheetName, sheetData);
                beanInfosFromXMLMap.put(sheetName, beanInfosFromXML);
                taskHeaderRowNums.put(sheetName, taskHeaderRowNum);
            }
            ProcessResult<Map<String, Object>> processResult = new ProcessResult<Map<String, Object>>(excelUtilLogger, sheetDatas);
            return processResult;
        } finally {
            workBook2007.close();
        }
    }

}
