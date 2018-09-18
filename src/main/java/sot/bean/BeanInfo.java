package sot.bean;

import java.util.ArrayList;
import java.util.List;

/**
 * Excel Bean 映射关系对象
 *
 * @author Monan
 * created on 2018/7/23 10:53
 */
public class BeanInfo {

    /**
     * 对象class名
     */
    private String className;

    /**
     * 是否允许插入时部分成功
     */
    private boolean isAllowFailOnInsert;

    /**
     * 错误限制
     */
    private int errorLimit = 0;


    /**
     * 错误计数器
     */
    private int errorCount = 0;

    /**
     * 对象属性名和对应位置list
     */
    private List<PropertyInfo> propertyInfos;

    /**
     * 添加属性信息
     * 多线程不安全
     *
     * @param propertyInfo Excel与Class属性映射关系对象
     */
    public void addPropertyInfo(PropertyInfo propertyInfo) {
        if (propertyInfos == null) {
            propertyInfos = new ArrayList<>();
        }
        propertyInfos.add(propertyInfo);
    }

    /**
     * Sets new 对象class名.
     *
     * @param className New value of 对象class名.
     */
    public void setClassName(String className) {
        this.className = className;
    }

    /**
     * Gets 对象class名.
     *
     * @return Value of 对象class名.
     */
    public String getClassName() {
        return className;
    }

    /**
     * Gets 是否允许插入时部分成功.
     *
     * @return Value of 是否允许插入时部分成功.
     */
    public boolean isIsAllowFailOnInsert() {
        return isAllowFailOnInsert;
    }

    /**
     * 增加一次错误计数
     */
    public void addError() {
        errorCount++;
    }

    /**
     * 判断是否已经超过错误次数限制
     *
     * @return {@code true} if errors is reached error limit,{@code false} if not;
     */
    public boolean isReachErrorLimit() {
        return (!isAllowFailOnInsert && errorCount >= errorLimit);
    }

    /**
     * Sets new 是否允许插入时部分成功.
     *
     * @param isAllowFailOnInsert New value of 是否允许插入时部分成功.
     */
    public void setIsAllowFailOnInsert(boolean isAllowFailOnInsert) {
        this.isAllowFailOnInsert = isAllowFailOnInsert;
    }


    /**
     * Gets 对象属性名和对应位置list.
     *
     * @return Value of 对象属性名和对应位置list.
     */
    public List<PropertyInfo> getPropertyInfos() {
        return propertyInfos;
    }

    /**
     * Sets new 对象属性名和对应位置list.
     *
     * @param propertyInfos New value of 对象属性名和对应位置list.
     */
    public void setPropertyInfos(List<PropertyInfo> propertyInfos) {
        this.propertyInfos = propertyInfos;
    }

    /**
     * Sets new 错误计数器.
     *
     * @param errorCount New value of 错误计数器.
     */
    public void setErrorCount(int errorCount) {
        this.errorCount = errorCount;
    }

    /**
     * Gets 错误计数器.
     *
     * @return Value of 错误计数器.
     */
    public int getErrorCount() {
        return errorCount;
    }

    /**
     * Gets 错误限制.
     *
     * @return Value of 错误限制.
     */
    public int getErrorLimit() {
        return errorLimit;
    }

    /**
     * Sets new 错误限制.
     *
     * @param errorLimit New value of 错误限制.
     */
    public void setErrorLimit(int errorLimit) {
        this.errorLimit = errorLimit;
    }


    @Override
    public String toString() {
        return "BeanInfo{" +
                "className='" + className + '\'' +
                ", isAllowFailOnInsert=" + isAllowFailOnInsert +
                ", propertyInfos=" + propertyInfos +
                '}';
    }
}
