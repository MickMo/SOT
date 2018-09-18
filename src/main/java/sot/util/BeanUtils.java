package sot.util;

import java.beans.BeanInfo;
import java.beans.IntrospectionException;
import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.HashMap;
import java.util.Map;

/**
 * <内容说明>
 *
 * @author Monan
 * created on 2018/9/3 10:19
 */
public final class BeanUtils {
    private BeanUtils() {}

    /**
     * Bean2Map <br>
     * 利用Introspector和PropertyDescriptor 将Bean --> Map
     *
     * @param obj Bean
     * @return Map
     */
    public static Map<String, Object> transBean2Map(Object obj) throws IllegalAccessException, IntrospectionException, InvocationTargetException {
        if (obj == null) {
            return null;
        }
        if(obj instanceof Map) {
            return (Map) obj;
        }
        Map<String, Object> map = new HashMap<String, Object>();
        try {
            BeanInfo beanInfo = Introspector.getBeanInfo(obj.getClass());
            PropertyDescriptor[] propertyDescriptors = beanInfo.getPropertyDescriptors();
            for (PropertyDescriptor property : propertyDescriptors) {
                String key = property.getName();
                // 过滤class属性
                if (!key.equals("class")) {
                    // 得到property对应的getter方法
                    Method getter = property.getReadMethod();
                    Object value = getter.invoke(obj);
                    map.put(key, value);
                }
            }
        } catch (Exception e) {
            throw e;
        }
        return map;
    }
}
