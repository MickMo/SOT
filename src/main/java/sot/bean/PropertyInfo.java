package sot.bean;

/**
 * 储存对应class的属性名和在excel中的位置
 *
 * @author Monan
 *         created on 2018/7/23 10:54
 */
public class PropertyInfo {

    /**
     * 对应excel列名
     */
    private String columnName;

    /**
     * 属性名
     */
    private String name;

    /**
     * 属性位置
     */
    private int index = -1;

    /**
     * 是否允许该数据为空
     */
    private boolean allowNull = false;

    /**
     * 属性不允许的值
     */
    private String[] notAllowValues;

    /**
     * 属性不允许的值 字符串
     */
    private String notAllowValuesString;

    /**
     * Minimum digit of property
     */
    private int minlength;

    /**
     * Minimum value of property
     */
    private int minValue;

    /**
     * 属性允许的值
     */
    private String[] allowValues;

    /**
     * 属性允许的值 字符串
     */
    private String allowValuesString;

    /**
     * Maximum digit of property
     */
    private int maxlength;

    /**
     * Maximum value of property
     */
    private int maxValue;

    /**
     * regex that use to check this property
     */
    private String regex;

    /**
     * Gets 属性名.
     *
     * @return Value of 属性名.
     */
    public String getName() {
        return name;
    }

    /**
     * Gets 属性位置.
     *
     * @return Value of 属性位置.
     */
    public int getIndex() {
        return index;
    }

    /**
     * Sets new 属性名.
     *
     * @param name New value of 属性名.
     */
    public void setName(String name) {
        this.name = name;
    }

    /**
     * Sets new 属性位置.
     *
     * @param index New value of 属性位置.
     */
    public void setIndex(int index) {
        this.index = index;
    }


    /**
     * Gets 是否允许该数据为空.
     *
     * @return Value of 是否允许该数据为空.
     */
    public boolean isAllowNull() {
        return allowNull;
    }

    /**
     * Sets new 是否允许该数据为空.
     *
     * @param isAllowNull New value of 是否允许该数据为空.
     */
    public void setIsAllowNull(boolean isAllowNull) {
        this.allowNull = isAllowNull;
    }

    /**
     * Gets 对应excel列名.
     *
     * @return Value of 对应excel列名.
     */
    public String getColumnName() {
        return columnName;
    }

    /**
     * Sets new 对应excel列名.
     *
     * @param columnName New value of 对应excel列名.
     */
    public void setColumnName(String columnName) {
        this.columnName = columnName;
    }


    /**
     * Sets new 属性允许的值.
     *
     * @param allowValues New value of 属性允许的值.
     */
    public void setAllowValues(String[] allowValues) {
        StringBuilder stringBuilder = new StringBuilder();
        stringBuilder.append("[");
        for (String allowValue : allowValues) {
            stringBuilder.append(allowValue + ",");
        }
        stringBuilder.deleteCharAt(stringBuilder.length() - 1);
        stringBuilder.append("]");
        this.allowValuesString = stringBuilder.toString();
        this.allowValues = allowValues;
    }


    /**
     * Gets 属性允许的值.
     *
     * @return Value of 属性允许的值.
     */
    public String[] getAllowValues() {
        return allowValues;
    }

    /**
     * Gets 属性允许的值 字符串.
     *
     * @return Value of 属性允许的值.
     */
    public String getAllowValuesString() {
        return allowValuesString;
    }


    /**
     * Gets Minimum digit of property.
     *
     * @return Value of Minimum digit of property.
     */
    public int getMinlength() {
        return minlength;
    }

    /**
     * Sets new Maximum digit of property.
     *
     * @param maxlength New value of Maximum digit of property.
     */
    public void setMaxlength(int maxlength) {
        this.maxlength = maxlength;
    }

    /**
     * Sets new 属性不允许的值.
     *
     * @param notAllowValues New value of 属性不允许的值.
     */
    public void setNotAllowValues(String[] notAllowValues) {
        StringBuilder stringBuilder = new StringBuilder();
        stringBuilder.append("[");
        for (String notAllowValue : notAllowValues) {
            stringBuilder.append(notAllowValue + ",");
        }
        stringBuilder.deleteCharAt(stringBuilder.length() - 1);
        stringBuilder.append("]");
        this.allowValuesString = stringBuilder.toString();
        this.notAllowValues = notAllowValues;
    }


    /**
     * Gets 属性允许的值 字符串.
     *
     * @return Value of 属性允许的值 字符串.
     */
    public String getNotAllowValuesString() {
        return notAllowValuesString;
    }

    /**
     * Gets regex that use to check this property.
     *
     * @return Value of regex that use to check this property.
     */
    public String getRegex() {
        return regex;
    }

    /**
     * Gets Maximum digit of property.
     *
     * @return Value of Maximum digit of property.
     */
    public int getMaxlength() {
        return maxlength;
    }

    /**
     * Sets new Minimum digit of property.
     *
     * @param minlength New value of Minimum digit of property.
     */
    public void setMinlength(int minlength) {
        this.minlength = minlength;
    }

    /**
     * Sets new Maximum value of property.
     *
     * @param maxValue New value of Maximum value of property.
     */
    public void setMaxValue(int maxValue) {
        this.maxValue = maxValue;
    }

    /**
     * Sets new Minimum value of property.
     *
     * @param minValue New value of Minimum value of property.
     */
    public void setMinValue(int minValue) {
        this.minValue = minValue;
    }

    /**
     * Gets Minimum value of property.
     *
     * @return Value of Minimum value of property.
     */
    public int getMinValue() {
        return minValue;
    }


    /**
     * Sets new regex that use to check this property.
     *
     * @param regex New value of regex that use to check this property.
     */
    public void setRegex(String regex) {
        this.regex = regex;
    }

    /**
     * Gets Maximum value of property.
     *
     * @return Value of Maximum value of property.
     */
    public int getMaxValue() {
        return maxValue;
    }

    /**
     * Gets 属性允许的值.
     *
     * @return Value of 属性允许的值.
     */
    public String[] getNotAllowValues() {
        return notAllowValues;
    }

}
