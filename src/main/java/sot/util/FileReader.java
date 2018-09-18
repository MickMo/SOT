package sot.util;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;

/**
 * 提供Jar包内/相对路径(绝对路径) 的文件读取功能
 * 基于application.properties 的 properties.location 配置
 *
 * @author Monan
 * created on 2018/8/17 11:03
 */

public final class FileReader {

    private FileReader() {
    }

    /**
     * 日志
     */
    private static Logger logger = LoggerFactory.getLogger(FileReader.class);

    /**
     * 读取文件
     * 根据配置自动决定文件位置
     *
     * @param path 文件位置
     * @return Properties
     */
    public static InputStream readFile(String path) {
        InputStream inputStream = readFileOutsideJarPackage(path);
        //读取配置文件
        if (inputStream != null) {
            //读取Jar包外配置文件
            return inputStream;

        } else {
            //读取Jar包内配置文件
            return readFileInsideJarPackage(path);
        }
    }


    /**
     * 读取Jar包内的properties文件
     *
     * @param path 文件位置
     * @return Properties
     */
    public static InputStream readFileInsideJarPackage(String path) {
        File file = null;
        InputStream in = null;
        try {
            file = new File(path);
            logger.info("Reading File inside Jar package {}", file.getPath());
            //读取Jar包内配置文件
            in = FileReader.class.getClassLoader().getResourceAsStream(path);
            if (in == null) {
                logger.warn("Reading file Inside the jar package.Resource " + path + " could not be read.");
                return null;
            }
            return in;
        } catch (Exception e) {
            e.printStackTrace();
            logger.warn("Could not read {} ,due bean {}", file.getAbsolutePath(), e.getMessage(), e);
            return null;
        }
    }

    /**
     * 读取Jar包内的properties文件
     *
     * @param path 文件位置
     * @return Properties
     */
    public static InputStream readFileOutsideJarPackage(String path) {
        File file = null;
        InputStream in = null;
        try {
            file = new File(path);
            logger.info("Reading file from {}", file.getAbsolutePath());

            //读取外部配置文件
            if (!file.exists()) {
                logger.warn("Reading file OutSide the jar package.Resource {} not exists.", file.getAbsolutePath());
                return null;
            }
            in = new FileInputStream(file);
            if (in == null) {
                logger.warn("Reading file OutSide the jar package.Resource " + path + " could not be read.");
                return null;
            }
            return in;
        } catch (Exception e) {
            e.printStackTrace();
            logger.warn("Could not read {} ,due bean {}", file.getAbsolutePath(), e.getMessage(), e);
            return null;
        }
    }


}
