package sot.util;

import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.io.SAXReader;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import sot.constant.SOTExcelConfigureConstant;
import sot.constant.SOTWordConfigureConstant;

import java.io.InputStream;

/**
 * XML文件读取工具
 *
 * @author Monan
 *         created on 2018/7/19 12:07
 */
public final class XmlReadUtil {
    private XmlReadUtil() {
    }

    /**
     * 日志
     */
    private static final Logger logger = LoggerFactory.getLogger(XmlReadUtil.class);


    /**
     * 读取 Excel导入任务字段映射文件
     *
     * @return EC接口配置
     */
    public static Document getMappingForExcelImportExportTask() {
        return readFromFile(SOTExcelConfigureConstant.CONFIG_FILE_PATH);
    }

    /**
     * 读取 Excel导入任务字段映射文件
     *
     * @return EC接口配置
     */
    public static Document getMappingForWordImportExportTask() {
        return readFromFile(SOTWordConfigureConstant.CONFIG_FILE_PATH);
    }

    /**
     * 读取XML文件
     *
     * @param filePath 文件路径
     * @return XML Document
     * @throws DocumentException XML读取错误
     */
    public static Document readFromFile(String filePath) {
        try {
            InputStream inputStream = FileReader.readFile(filePath);
            SAXReader reader = new SAXReader();
            Document document = reader.read(inputStream);
            return document;
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    /**
     * 读取XML文件
     *
     * @param inputStream xml文件流
     * @return XML Document
     * @throws DocumentException XML读取错误
     */
    public static Document readFromFile(InputStream inputStream) throws DocumentException {
        try {
            SAXReader reader = new SAXReader();
            Document document = reader.read(inputStream);
            return document;
        } catch (DocumentException e) {
            e.printStackTrace();
            throw e;
        }
    }


}
