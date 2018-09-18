package sot.util;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

/**
 * 配置文件读取器
 *
 * @author zrhong
 * created on 2018/6/7 15:12
 */
public final class PropertiesReader {

    private PropertiesReader() {
    }

    /**
     * 日志
     */
    private static Logger logger = LoggerFactory.getLogger(PropertiesReader.class);

    /**
     * 读取properties文件
     * 根据配置自动决定文件位置
     *
     * @param path 文件位置
     * @return Properties
     */
    public static Properties getPropInstance(String path) {
        InputStream inputStream = FileReader.readFile(path);
        return loadProperties(inputStream);
    }

    /**
     * 读取Jar包内的properties文件
     *
     * @param path 文件位置
     * @return Properties
     */
    public static Properties getPropInstanceInsideJarPackage(String path) {
        InputStream inputStream = FileReader.readFileInsideJarPackage(path);
        return loadProperties(inputStream);
    }

    /**
     * 读取Jar包内的properties文件
     *
     * @param path 文件位置
     * @return Properties
     */
    public static Properties getPropInstanceOutsideJarPackage(String path) {
        InputStream inputStream = FileReader.readFileOutsideJarPackage(path);
        return loadProperties(inputStream);
    }


    /**
     * 装载Properties
     *
     * @param inputStream 文件输入流
     * @return Properties
     */
    private static Properties loadProperties(InputStream inputStream) {
        Properties properties = new Properties();
        try {
            properties.load(inputStream);
            return properties;
        } catch (Exception e) {
            logger.warn("Could not load properties file,due bean {}", e.getMessage());
            return null;
        } finally {
            if (null != inputStream) {
                try {
                    inputStream.close();
                } catch (IOException e) {
                    logger.warn("Could not close inputStream,due bean {}", e.getMessage(), e);
                }
            }
        }
    }

}
