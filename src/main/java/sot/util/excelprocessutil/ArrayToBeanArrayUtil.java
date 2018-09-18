package sot.util.excelprocessutil;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import sot.bean.BeanInfo;
import sot.bean.PropertyInfo;
import sot.bean.exception.ExcelProcessingException;
import sot.wordutil.logger.ExcelUtilLogger;

import java.lang.reflect.Method;
import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Excel数据转Object工具类
 *
 * @author Monan
 *         created on 2018/7/19 12:07
 */
@SuppressWarnings("Duplicates")
public final class ArrayToBeanArrayUtil {

    private ArrayToBeanArrayUtil() {

    }


    /**
     * This is Shit......
     * 通过BeanInfo中的映射关系,反射生成对应的Class并填充数据
     *
     * @param sheet            excel sheet数据
     * @param headerRowNum     表头行位置
     * @param beanInfosFromXML Excel Bean 映射关系对象
     * @param excelUtilLogger  错误日志记录器
     * @return POJO类的二维数组, 第二维数组中可能存在多个类型的POJO
     * @throws ExcelProcessingException 反射错误 or 数据校验错误
     */
    public static List<List<Object>> data2DarrayTo2DBeanArray(Sheet sheet, int headerRowNum, List<BeanInfo> beanInfosFromXML, ExcelUtilLogger excelUtilLogger)
            throws
            ExcelProcessingException {
        List<List<Object>> dataList = new ArrayList<>();
        StringBuilder errorInfoBuilder = new StringBuilder();
        //创建Bean
        int rowNum = (headerRowNum + 1);
        for (; rowNum <= sheet.getLastRowNum(); rowNum++) {
            Row row = sheet.getRow(rowNum);
            //空行判断
            if (emptyRowJudgement(row)) {
                continue;
            }
            ArrayList<Object> beanList = new ArrayList<>();
            HashMap tempMap = new HashMap();
            for (BeanInfo beanInfo : beanInfosFromXML) {
                String className = beanInfo.getClassName();

                HashMap<String, Object> propertyMap = new HashMap<>();
                //格式化数据
                List<PropertyInfo> propertyInfos = beanInfo.getPropertyInfos();
                for (PropertyInfo propertyInfo : propertyInfos) {
                    int index = propertyInfo.getIndex();
                    if (index < 0) {
                        //没有该属性的映射位置
                        //错误信息
                        errorInfoBuilder.setLength(0);
                        errorInfoBuilder.append("Could not found Property [");
                        errorInfoBuilder.append(propertyInfo.getColumnName());
                        errorInfoBuilder.append("] ");
                        errorInfoBuilder.append("on row:" + (rowNum + 1));

                        RuntimeException runtimeException = new RuntimeException(errorInfoBuilder.toString());
                        excelUtilLogger.recordExcelException(beanInfo, row, runtimeException, !propertyInfo.isAllowNull());
                        continue;
                    }


                    Cell cell = row.getCell(index);
                    //在Excel表格第一行中没有找到对应的属性名,检查是否允许该属性为空
                    if (cell == null) {
                        //错误信息
                        errorInfoBuilder.setLength(0);
                        errorInfoBuilder.append("Property [");
                        errorInfoBuilder.append(propertyInfo.getColumnName());
                        errorInfoBuilder.append("]'s value doesn't exist");
                        errorInfoBuilder.append(" at index:" + (index + 1));
                        errorInfoBuilder.append(",on row:" + (rowNum + 1));

                        RuntimeException runtimeException = new RuntimeException(errorInfoBuilder.toString());
                        excelUtilLogger.recordExcelException(beanInfo, row, runtimeException, !propertyInfo.isAllowNull());
                        continue;
                    }


                    Object valueObject = ConversionUtil.transformCell(cell);
//                    String cellValueAsString = ConversionUtil.getCellValueAsString(cell);
                    //检查是值转换是否成功,并且是否允许该属性为空
                    if (valueObject == null || StringUtils.isEmpty(valueObject.toString())) {
                        //错误信息
                        errorInfoBuilder.setLength(0);
                        errorInfoBuilder.append("Could not convert Property [");
                        errorInfoBuilder.append(propertyInfo.getColumnName());
                        errorInfoBuilder.append("]'s value ");
                        errorInfoBuilder.append(",at index:" + (index + 1));
                        errorInfoBuilder.append(",on row:" + (rowNum + 1));
                        errorInfoBuilder.append(",with Cell value:" + ConversionUtil.getCellValueAsString(cell));
                        errorInfoBuilder.append(",and type:" + ConversionUtil.getCellTypeAsString(cell));

                        RuntimeException runtimeException = new RuntimeException(errorInfoBuilder.toString());
                        excelUtilLogger.recordExcelException(beanInfo, row, runtimeException, !propertyInfo.isAllowNull());
                        continue;
                    }

                    //检查是否为允许的值
                    String[] allowValues = propertyInfo.getAllowValues();
                    if (allowValues != null && allowValues.length > 0) {
                        String valueStr = valueObject + "";
                        boolean isValidAllowValue = false;
                        for (String allowValue : allowValues) {
                            if (allowValue.equals(valueStr)) {
                                isValidAllowValue = true;
                                break;
                            }
                        }
                        if (!isValidAllowValue) {
                            //错误信息
                            errorInfoBuilder.setLength(0);
                            errorInfoBuilder.append("Value '" + valueStr + "'");
                            errorInfoBuilder.append(",at index:" + (index + 1));
                            errorInfoBuilder.append(",on row:" + (rowNum + 1));
                            errorInfoBuilder.append(",with Cell value:" + ConversionUtil.getCellValueAsString(cell));
                            errorInfoBuilder.append(",and type:" + ConversionUtil.getCellTypeAsString(cell));

                            errorInfoBuilder.append(",is not allow for Property [");
                            errorInfoBuilder.append(propertyInfo.getColumnName());
                            errorInfoBuilder.append("],allow values are " + propertyInfo.getAllowValuesString());

                            RuntimeException runtimeException = new RuntimeException(errorInfoBuilder.toString());
                            excelUtilLogger.recordExcelException(beanInfo, row, runtimeException, !propertyInfo.isAllowNull());
                            continue;
                        }
                    }


                    propertyMap.put(propertyInfo.getName(), valueObject);
                }
                //创建Bean
                try {
                    if (StringUtils.isNotBlank(className)) {
                        Object object = ArrayToBeanArrayUtil.creatObjectFromBeanData(className, propertyMap);
                        beanList.add(object);
                    } else {
                        tempMap.putAll(propertyMap);
                    }
                } catch (Exception e) {
                    excelUtilLogger.recordExcelError(beanInfo, row, e);
                }
            }
            if (tempMap.size()>0) {
                beanList.add(tempMap);
            }
            dataList.add(beanList);
        }
        return dataList;
    }


    /**
     * 空行检查
     *
     * @param row 行数据
     * @return 是否为空行
     */
    private static boolean emptyRowJudgement(Row row) {
        StringBuilder stringBuilder = new StringBuilder();
        for (Cell cell : row) {
            String cellValueAsString = ConversionUtil.getCellValueAsString(cell);
            if (StringUtils.isNotEmpty(cellValueAsString)) {
                stringBuilder.append(cellValueAsString.trim());
            }
        }
        return StringUtils.isEmpty(stringBuilder.toString());
    }

    /**
     * 从beanData中获取数据,生成Obj对象
     *
     * @param className 包含包信息的类名
     * @param beanData  对象数据,属性名-数据Map
     * @return 已填充数据的对象
     * @throws Exception                反射异常
     * @throws IllegalArgumentException 数据转换异常
     */
    private static Object creatObjectFromBeanData(String className, Map<String, Object> beanData) throws Exception, IllegalArgumentException {
        Class<?> clazz = Class.forName(className);
        // 获取obj引用的对象字节码
        Object tempObj = clazz.newInstance();
        Method[] declaredMethods = clazz.getMethods();

        Object paramaterValue = null;
        Object prasedParameter = null;

        // 找合适的set方法,转换到合适的参数类型,执行反射
        for (String fieldName : beanData.keySet()) {
            paramaterValue = beanData.get(fieldName);
            // 对象数据为空,跳过
            if (paramaterValue == null) {
                continue;
            }
            for (Method method : declaredMethods) {
                Class<?>[] parameterTypes = method.getParameterTypes();
                // 跳过不适合的方法
                if (!generalCheckBeforeInvoking(method, fieldName)) {
                    continue;
                }
                // 字符串方法
                if (parameterTypes[0].getSimpleName().equals(String.class.getSimpleName())) {
                    method.invoke(tempObj, paramaterValue);
                } else {
                    // 其他参数类型方法
                    prasedParameter = reflectionParameterTypeHandller(parameterTypes[0].getSimpleName(), paramaterValue);
                    method.invoke(tempObj, prasedParameter);
                }
            }
        }
        return tempObj;
    }


    /**
     * 对方法的综合检查
     *
     * @param method    Bean中的方法
     * @param fieldName 目标字段名
     * @return 检查结果 <br> @code{true} means pass,otherwise the checking had failed..
     */
    private static boolean generalCheckBeforeInvoking(Method method, String fieldName) {
        String methodName = method.getName();
        Class<?>[] parameterTypes = method.getParameterTypes();
        // 跳过多参数方法
        if (parameterTypes.length > 1) {
            return false;
        }
        // 跳过非set的方法
        if (!methodName.toLowerCase().contains("set")) {
            return false;
        }
        // 寻找合适的方法setFsfsd
        if (methodName.toLowerCase().replace("set", "").equals(fieldName.toLowerCase())) {
            return true;
        }
        return false;
    }


    /**
     * 反射参数自动转换 <br/>
     * 支持:String --> Timestamp,Double,Long,Integer,
     *
     * @param targetParameterType 参数类型
     * @param value               数据
     * @return 对应类型的参数对象
     * @throws IllegalArgumentException 错误的数据-类型搭配
     */
    private static Object reflectionParameterTypeHandller(String targetParameterType, Object value) throws IllegalArgumentException {
        if (value == null || StringUtils.isEmpty(value.toString())) {
            throw new IllegalArgumentException("Could not converse null or empty Value!");
        }
        if (targetParameterType.equals(Timestamp.class.getSimpleName())) {
            // Timestamp
            return ConversionUtil.praseDate(value.toString());
        } else if (targetParameterType.equals(Double.class.getSimpleName())) {
            // Double
            return Double.parseDouble(value.toString());
        } else if (targetParameterType.equals(int.class.getSimpleName())
                || targetParameterType.equals(Integer.class.getSimpleName())) {
            // int or Integer
            return Integer.parseInt(value.toString());
        } else if (targetParameterType.equals(Long.class.getSimpleName())) {
            // Long
            return Long.parseLong(value.toString());
        } else if (targetParameterType.equals(Boolean.class.getSimpleName())) {
            // Boolean
            switch (value.toString().toLowerCase()) {
                case "false":
                    return false;
                case "true":
                    return true;
                default:
                    throw new IllegalArgumentException("Unsupported Boolean Value:" + value);
            }
        } else {
            throw new IllegalArgumentException("Unsupported Parameter Type:" + targetParameterType);
        }
    }
}
