package sot.util.fileutil;

/**
 * 创建XML SQL 工具类
 *
 * @author Monan
 * created on 2018/7/25 16:36
 */
public final class XMLSQLUtil {

    private XMLSQLUtil() {
    }

    /**
     * XML SQL Prefix Part "sql[@"
     * [@'ParameterName'='ParameterValue']
     */
    private static String sqlPrefix = "[@";
    /**
     * XML SQL Middle Part: "='"
     * [@'ParameterName'='ParameterValue']
     */
    private static String sqlMiddle = "='";
    /**
     * XML SQL Suffix Part: "']"
     * [@'ParameterName'='ParameterValue']
     */
    private static String sqlSuffix = "']";

    /**
     * 创建XML SQL
     *
     * @param name  ParameterName
     * @param value ParameterValue
     * @return sql[@'ParameterName'='ParameterValue']
     */
    public static String generatXMLSQL(String name, String value) {
        StringBuilder stringBuilder = new StringBuilder();
        stringBuilder.append(sqlPrefix);
        stringBuilder.append(name);
        stringBuilder.append(sqlMiddle);
        stringBuilder.append(value);
        stringBuilder.append(sqlSuffix);
        return stringBuilder.toString();
    }
}
