package sot.util.fileutil;

import org.apache.commons.lang3.StringUtils;
import org.dom4j.Document;
import org.dom4j.Node;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import sot.bean.BeanInfo;
import sot.bean.PropertyInfo;
import sot.bean.exception.ExcelProcessingException;
import sot.constant.SOTExcelConfigureConstant;
import sot.util.ClassFinder;
import sot.wordutil.logger.ExcelUtilLogger;

import java.util.ArrayList;
import java.util.List;

/**
 * 读取Excel Bean映射关系工具类
 *
 * @author Monan
 *         created on 2018/7/26 12:07
 */
public final class ReadExcelMappingUtil {
    private ReadExcelMappingUtil() {
    }

    /**
     * 日志
     */
    @SuppressWarnings({"CheckStyle"})
    private static final Logger logger = LoggerFactory.getLogger(ReadExcelMappingUtil.class);

    /**
     * 获取任务描述
     *
     * @param taskNode 任务节点
     * @return 任务描述
     */
    public static String getTaskComment(Node taskNode) {
        return taskNode.valueOf("@" + SOTExcelConfigureConstant.TASK_COMMENT);
    }

    /**
     * 获取表头行数
     *
     * @param taskNode 任务节点
     * @return 表头行数
     */
    public static int getTaskHeaderRowNum(Node taskNode, String sheetName) {
        //TODO 需要针对Workbook增加 SheetName的处理
        String taskHeaderRowNumStr = taskNode.valueOf("@" + SOTExcelConfigureConstant.TASK_HEADERROWNUM);
        if (StringUtils.isEmpty(taskHeaderRowNumStr)) {
            taskHeaderRowNumStr = "0";
        } else {
            taskHeaderRowNumStr = (Integer.valueOf(taskHeaderRowNumStr) - 1) + "";
        }
        return Integer.valueOf(taskHeaderRowNumStr);
    }


    /**
     * 从Bean Excel映射关系 xml文档中获取指定任务名的任务节点
     *
     * @param document Bean Excel映射关系 xml文档对象
     * @param taskName 任务名
     * @return 任务节点
     * @throws RuntimeException 找不到目标任务
     * @throws Exception        XML Mapping文件读取错误
     */
    public static Node getTaskNode(Document document, String taskName) throws RuntimeException, Exception {
        logger.info("Searching Excel Mapping task with task name:{}...", taskName);
        String xmlSQL = XMLSQLUtil.generatXMLSQL(SOTExcelConfigureConstant.TASK_NAME, taskName);
        try {
            Node taskNode = document.selectSingleNode("/root/task" + xmlSQL);
            if (taskNode == null) {
                //找不到目标任务
                logger.error("No task found in Excel Mapping Config File:{}, with taskName:{}", SOTExcelConfigureConstant
                        .CONFIG_FILE_PATH, taskName);
                throw new RuntimeException("No task found in Excel Mapping Config with taskName '" + taskName + "'");
            }
            return taskNode;
        } catch (Exception e) {
            //XML Mapping文件读取错误
            logger.error("Could not Read XML Mapping File with SQL:{}", xmlSQL);
            throw e;
        }
    }


    /**
     * 根据Bean Excel映射关系XML文件,生成BeanInfo对象数组
     *
     * @param taskNode        excel映射文件中的class节点
     * @param excelUtilLogger 错误记录器
     * @return Bean Excel映射关系对象 数组
     * @throws ExcelProcessingException XML文件解析错误
     */
    public static List<BeanInfo> getBeanInfosFromXML(Node taskNode, ExcelUtilLogger excelUtilLogger, String sheetName) throws
            ExcelProcessingException {
        logger.info("Start Converting XML Excel Mapping Info bean BeanInfo...");
        //遍历sheet
        List<Node> classList = taskNode.selectNodes(SOTExcelConfigureConstant.CLASS_ELEMENT_NAME);
        if (classList.size() < 1) {
            return null;
        }
        List<BeanInfo> beanInfos = new ArrayList<>();
        //遍历class
        for (Node classNode : classList) {
            BeanInfo beanInfo = new BeanInfo();

            //isAllowFailonInsert:是否允许部分属性为null,默认为false
            boolean isAllowFailOnInsert = Boolean.valueOf(classNode.valueOf("@" + SOTExcelConfigureConstant.CLASS_IS_ALLOW_FAIL_ON_INSERT));
            beanInfo.setIsAllowFailOnInsert(isAllowFailOnInsert);
            if (!isAllowFailOnInsert) {
                //insertFailErrorLimit:允许最大的错误次数
                String insertFailErrorLimit = classNode.valueOf("@" + SOTExcelConfigureConstant.CLASS_INSERT_FAIL_ERROR_LIMIT);
                if (StringUtils.isNotEmpty(insertFailErrorLimit)) {
                    beanInfo.setErrorLimit(Integer.parseInt(insertFailErrorLimit));
                }
            }
            try {
                List<PropertyInfo> propertyInfos = generatePropertyInfo(classNode);
                beanInfo.setPropertyInfos(propertyInfos);
                beanInfos.add(beanInfo);
                logger.debug("BeanInfo generated with[isAllowFailOnInsert{},propertyMappingCount:{}]", isAllowFailOnInsert, beanInfo
                        .getPropertyInfos().size());
            } catch (Exception e) {
                excelUtilLogger.recordXMLError(e);
            }
        }
        return beanInfos;
    }

    /**
     * 根据Bean Excel映射关系XML文件,生成BeanInfo对象数组
     *
     * @param taskNode        excel映射文件中的class节点
     * @param excelUtilLogger 错误记录器
     * @return Bean Excel映射关系对象 数组
     * @throws ExcelProcessingException XML文件解析错误
     */
    public static List<BeanInfo> getBeanInfosFromXML(Node taskNode, String[] packages, ExcelUtilLogger excelUtilLogger, String sheetName) throws
            ExcelProcessingException {
        logger.info("Start Converting XML Excel Mapping Info bean BeanInfo...");
        //遍历sheet
        List<Node> classList = taskNode.selectNodes(SOTExcelConfigureConstant.CLASS_ELEMENT_NAME);
        if (classList.size() < 1) {
            return null;
        }
        List<BeanInfo> beanInfos = new ArrayList<>();
        //遍历class
        for (Node classNode : classList) {
            BeanInfo beanInfo = new BeanInfo();

            //处理Class类名
            String clazzName = classNode.valueOf("@" + SOTExcelConfigureConstant.CLASS_NAME);
//            if (StringUtils.isBlank(clazzName)) {
//
//            } else {
//
//            }
            if (clazzName.lastIndexOf(".") < 0) {
                logger.info("Classname does not contain full packages info,generating...");
                //没有完整类名
                try {
                    clazzName = ClassFinder.findClassBySimpleName(clazzName, packages).getName();
                } catch (ClassNotFoundException e) {
                    logger.error(e.getMessage());
                    excelUtilLogger.recordXMLError(e);
                }
            }
            beanInfo.setClassName(clazzName);

            //isAllowFailonInsert:是否允许部分属性为null,默认为false
            boolean isAllowFailOnInsert = Boolean.valueOf(classNode.valueOf("@" + SOTExcelConfigureConstant.CLASS_IS_ALLOW_FAIL_ON_INSERT));
            beanInfo.setIsAllowFailOnInsert(isAllowFailOnInsert);
            if (!isAllowFailOnInsert) {
                //insertFailErrorLimit:允许最大的错误次数
                String insertFailErrorLimit = classNode.valueOf("@" + SOTExcelConfigureConstant.CLASS_INSERT_FAIL_ERROR_LIMIT);
                if (StringUtils.isNotEmpty(insertFailErrorLimit)) {
                    beanInfo.setErrorLimit(Integer.parseInt(insertFailErrorLimit));
                }
            }
            try {
                List<PropertyInfo> propertyInfos = generatePropertyInfo(classNode);
                beanInfo.setPropertyInfos(propertyInfos);
                beanInfos.add(beanInfo);
                logger.debug("BeanInfo generated with[ className:{},isAllowFailOnInsert{},propertyMappingCount:{}", clazzName, isAllowFailOnInsert, beanInfo
                        .getPropertyInfos().size());
            } catch (Exception e) {
                excelUtilLogger.recordXMLError(e);
            }
        }
        return beanInfos;
    }


    /**
     * 生成 List<PropertyInfo>
     *
     * @param classNode excel映射文件中的class节点
     * @return Excel映射Class属性关系List
     * @throws RuntimeException XML文件错误,class节点解析错误
     */
    private static List<PropertyInfo> generatePropertyInfo(Node classNode) throws RuntimeException {
        //遍历columns
        List<Node> columns = classNode.selectNodes(SOTExcelConfigureConstant.COLUMNS_ELEMENT_NAME);
        if (columns.size() < 1) {
            return null;
        }
        List<PropertyInfo> properties = new ArrayList<>();
        for (Node column : columns) {
            PropertyInfo propertyInfo = new PropertyInfo();

            String columnName = column.valueOf("@" + SOTExcelConfigureConstant.COLUMN_NAME);
            String propertyName = column.valueOf("@" + SOTExcelConfigureConstant.PROPERTY_NAME);
            String isAllowNullStr = column.valueOf("@" + SOTExcelConfigureConstant.IS_ALLOW_PROPERTY_NULL);
            String minLengthStr = column.valueOf("@" + SOTExcelConfigureConstant.MIN_LENGTH);
            String minValueStr = column.valueOf("@" + SOTExcelConfigureConstant.MIN_VALUE);
            String maxLengthStr = column.valueOf("@" + SOTExcelConfigureConstant.MAX_LENGTH);
            String maxValueStr = column.valueOf("@" + SOTExcelConfigureConstant.MAX_VALUE);
            String regexStr = column.valueOf("@" + SOTExcelConfigureConstant.REGEX);

            if (StringUtils.isNotEmpty(columnName) && StringUtils.isNotEmpty(propertyName)) {
                propertyInfo.setColumnName(columnName);
                propertyInfo.setName(propertyName);
                if (StringUtils.isNumeric(minLengthStr)) {
                    propertyInfo.setMinlength(Integer.valueOf(minLengthStr));
                }
                if (StringUtils.isNumeric(minValueStr)) {
                    propertyInfo.setMinValue(Integer.valueOf(minValueStr));
                }

                if (StringUtils.isNumeric(maxLengthStr)) {
                    propertyInfo.setMaxlength(Integer.valueOf(maxLengthStr));
                }
                if (StringUtils.isNumeric(maxValueStr)) {
                    propertyInfo.setMaxValue(Integer.valueOf(maxValueStr));
                }
                if (StringUtils.isNotBlank(regexStr)) {
                    propertyInfo.setRegex(regexStr);
                }
                if (StringUtils.isNotEmpty(isAllowNullStr)) {
                    propertyInfo.setIsAllowNull(Boolean.valueOf(isAllowNullStr));
                }

                //处理允许值
                //insertFailErrorLimit:允许最大的错误次数
                String allowValuesStr = column.valueOf("@" + SOTExcelConfigureConstant.ALLOW_VALUE);
                if (StringUtils.isNotEmpty(allowValuesStr)) {
                    String[] allowValues = allowValuesStr.split(",");
                    propertyInfo.setAllowValues(allowValues);
                }
                //处理不允许值
                //insertFailErrorLimit:允许最大的错误次数
                String notAllowValuesStr = column.valueOf("@" + SOTExcelConfigureConstant.NOT_ALLOW_VALUE);
                if (StringUtils.isNotEmpty(notAllowValuesStr)) {
                    String[] notAllowValues = notAllowValuesStr.split(",");
                    propertyInfo.setNotAllowValues(notAllowValues);
                }

                logger.debug("PropertyInfo generated with[ columnName:{},classPropertyName:{},isAllowNull:{}", columnName, propertyName, propertyInfo
                        .isAllowNull());
                properties.add(propertyInfo);
                continue;
            } else {
                //XML文件错误
                throw new RuntimeException("One or more attribute(s) are missing under the " + column.getName() + "node");
            }

        }
        return properties;
    }
}
