package sot.util;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.HashMap;
import java.util.Map;
import java.util.Properties;

/**
 * 配置文件读取器
 *
 * @author zrhong
 * created on 2018/6/7 15:12
 */
public final class ConfigReader {

    private ConfigReader() {
    }

    /**
     * 配置文件缓存
     */
    private static Map<String, Properties> mappingConfig = new HashMap<>();

    /**
     * 日志
     */
    private static Logger logger = LoggerFactory.getLogger(ConfigReader.class);


    /**
     * 读取Excel导出的配置文件
     *
     * @return Excel导出配置
     */
    public static Properties getMappingForExcelExportInstance() {
        return getMappingPropInstance("config/excelImportExport.properties");
    }


    /**
     * 读取properties文件
     *
     * @param path 文件位置
     * @return Properties
     */
    public static Properties getMappingPropInstance(String path) {
        if (mappingConfig.containsKey(path)) {
            return mappingConfig.get(path);
        }
        Properties propInstance = PropertiesReader.getPropInstance(path);
        if (propInstance != null) {
            mappingConfig.put(path, propInstance);
        }
        return mappingConfig.get(path);
    }


    /**
     * 清空所有配置缓存
     */
    public static void resetAll() {
        mappingConfig.clear();
    }


}
