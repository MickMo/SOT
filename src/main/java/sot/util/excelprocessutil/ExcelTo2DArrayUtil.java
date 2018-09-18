package sot.util.excelprocessutil;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import sot.bean.exception.ExcelProcessingException;
import sot.wordutil.logger.ExcelUtilLogger;
import sot.bean.BeanInfo;
import sot.bean.PropertyInfo;

import java.util.ArrayList;
import java.util.List;

/**
 * Excel数据转换工具类
 *
 * @author Monan
 * created on 2018/7/19 12:06
 */
public final class ExcelTo2DArrayUtil {
    private ExcelTo2DArrayUtil() {
    }


    /**
     * 表格数据转换到BeanInfo 数组
     *
     * @param sheet            表格sheet
     * @param headerRowNum     表头行位置
     * @param beanInfosFromXML Excel Bean 映射关系对象
     * @param excelUtilLogger  错误记录器
     * @throws ExcelProcessingException 数据处理相关的异常
     */
    public static void sheetToBeanInfoArray(Sheet sheet, int headerRowNum, List<BeanInfo> beanInfosFromXML, ExcelUtilLogger excelUtilLogger) throws
            ExcelProcessingException {
        try {
            //读取Excel文件开头行,用于获取属性位置
            Row firstRow = sheet.getRow(headerRowNum);
            List<String> columnNames = new ArrayList<>();
            for (Cell cell : firstRow) {
                columnNames.add(ConversionUtil.getCellValueAsString(cell));
            }
            //保存属性位置
            for (BeanInfo beanInfo : beanInfosFromXML) {
                recordPropertyIndex(beanInfo.getPropertyInfos(), columnNames);
            }
        } catch (Exception e) {
            excelUtilLogger.recordXMLError(e);
        }
    }


    /**
     * 记录指定属性在Excel表中的位置
     *
     * @param propertyInfos 属性对象
     * @param columnNames   excel表格的第一行对应的数据
     * @throws RuntimeException 没有找到对应的属性名异常
     */
    private static void recordPropertyIndex(List<PropertyInfo> propertyInfos, List<String> columnNames) throws RuntimeException {
        for (PropertyInfo propertyInfo : propertyInfos) {
            //查找对应的属性名在excel表格第一行中的位置
            boolean hasFoundPropertyIndex = false;
            for (int i = 0; i < columnNames.size(); i++) {
                String propertyMappingName = propertyInfo.getColumnName().trim().toLowerCase();
                String columnName = columnNames.get(i);
                if (StringUtils.isEmpty(columnName)) {
                    continue;
                }
                columnName = columnName.trim().toLowerCase();

                hasFoundPropertyIndex = propertyMappingName.equals(columnName);
                if (hasFoundPropertyIndex) {
                    //保存位置
                    propertyInfo.setIndex(i);
                    break;
                }
            }
            //在Excel表格第一行中没有找到对应的属性名,检查是否允许该属性为空
            if (!hasFoundPropertyIndex && !propertyInfo.isAllowNull()) {
                throw new RuntimeException("No property[" + propertyInfo.getColumnName() + "] found in Excel's first row!");
            }
        }
    }
}
