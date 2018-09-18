package sot.util;

import org.apache.commons.lang3.ArrayUtils;

/**
 * Class Finder Utils
 *
 * @author Monan
 * created on 2018/7/26 11:56
 */
public final class ClassFinder {
    private ClassFinder() {
    }
    private final static String CLASS_SUFFIX = ".class";

    /**
     * Find class Object base on full reference or package locations store in array
     * @param simpleClassName Simple name of a class
     * @param targets         A String array which storing package locations of bean or full reference of a bean
     * @return target Class Object
     * @throws ClassNotFoundException Could not find target class.
     */
    public static Class<?> findClassBySimpleName(String simpleClassName, String[] targets) throws ClassNotFoundException {
        for (int i = 0; i < targets.length; i++) {
            try {
                String target = targets[i];
                if (target.toLowerCase().endsWith(CLASS_SUFFIX) && target.toLowerCase().contains(simpleClassName.toLowerCase())){
                    return Class.forName(target);
                } else {
                    return Class.forName(target + "." + simpleClassName);
                }
            } catch (ClassNotFoundException e) {
                continue;
            }
        }
        //nothing found: return null or throw ClassNotFoundException
        throw new ClassNotFoundException("No Class found under:" + ArrayUtils.toString(targets) + ", for Simple Class Name:" + simpleClassName);
    }


}
