package sot.constant;

import java.io.File;

/**
 * 基本导入配置常量
 *
 * @author Monan
 *         created on 2018/7/24 15:05
 */
public class BasicSOTConfigureConstant {
    protected BasicSOTConfigureConstant() {
    }

    /**
     * Excel映射XML文件位置
     */
    public static final String CONFIG_FILE_PATH = "./" + File.separator + "config" + File.separator + "SOTConfig.xml";

    /**
     * 根节点名字 配置名
     */
    public static final String ROOT_ELEMENT_NAME = "root";


    //    任务配置
    /**
     * task节点名字 配置名
     */
    public static final String TASK_ELEMENT_NAME = "task";
    /**
     * 任务名 配置名
     */
    public static final String TASK_NAME = "taskName";
    /**
     * 任务说明 配置名
     */
    public static final String TASK_COMMENT = "comment";
    /**
     * 指定表头行 配置名
     */
    public static final String TASK_HEADERROWNUM = "headerRowNum";


    //    class配置
    /**
     * class节点名字 配置名
     */
    public static final String CLASS_ELEMENT_NAME = "class";
    /**
     * 是否允许插入失败  配置名
     */
    public static final String CLASS_IS_ALLOW_FAIL_ON_INSERT = "isAllowFailonInsert";
    /**
     * 插入失败限制,超过限制将立即终止所有数据插入操作
     */
    public static final String CLASS_INSERT_FAIL_ERROR_LIMIT = "insertFailErrorLimit";


    //    属性映射配置
    /**
     * columns节点名字 配置名
     */
    public static final String COLUMNS_ELEMENT_NAME = "columns";
    /**
     * 表格列名  配置名
     */
    public static final String COLUMN_NAME = "columnName";
    /**
     * 类属性名  配置名
     */
    public static final String PROPERTY_NAME = "propertyName";
    /**
     * 允许该字段为空  配置名
     */
    public static final String IS_ALLOW_PROPERTY_NULL = "isAllowNull";
    /**
     * 字段允许值  配置名
     */
    public static final String ALLOW_VALUE = "allowValue";

    /**
     * 字段最大长度  配置名
     */
    public static final String MAX_LENGTH = "maxLength";

    /**
     * 字段最大值  配置名
     */
    public static final String MAX_VALUE = "maxValue";
    /**
     * 字段允许值  配置名
     */
    public static final String NOT_ALLOW_VALUE = "notAllowValue";
    /**
     * 字段最大长度  配置名
     */
    public static final String MIN_LENGTH = "minLength";

    /**
     * 字段最大值  配置名
     */
    public static final String MIN_VALUE = "minValue";

    /**
     * 字段最大值  配置名
     */
    public static final String REGEX = "regex";
}
