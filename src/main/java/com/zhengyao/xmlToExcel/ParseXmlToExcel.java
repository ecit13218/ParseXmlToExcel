package com.zhengyao.xmlToExcel;

/**
 * Created by 郑瑶 on 2017/8/14.
 */
public class ParseXmlToExcel {
    public static void main(String[] args) {
        ParseXmlToExcelImpl parseXmlToExcelImpl = new ParseXmlToExcelImpl(XmlVo.class, "G:\\F0.xml", "G:\\test.xls");
        parseXmlToExcelImpl.parseAndExport();
    }
}
