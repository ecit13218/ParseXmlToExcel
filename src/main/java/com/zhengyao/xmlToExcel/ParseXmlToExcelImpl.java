package com.zhengyao.xmlToExcel;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.dom4j.Attribute;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.*;

/**
 * Created by 郑瑶 on 2017/8/14.
 */
public class ParseXmlToExcelImpl {
    private static Field[] fields;
    private static List<Object> xmlVoList = new ArrayList<Object>();
    private static Object xmlVo;
    private static Class clazz;
    private static Map<String, Class> fieldString = new HashMap<String, Class>();
    private static int count = 0;
    private static String xmlPath;
    private static String excelPath;

    public ParseXmlToExcelImpl(Class clazz,String xmlPath,String excelPath) {
        this.clazz=clazz;
        this.xmlPath=xmlPath;
        this.excelPath=excelPath;
    }
    public void parseAndExport(){
        HSSFWorkbook wb = null;
        FileOutputStream fileOutputStream = null;
        try {
            wb = wbSetPrepare();
            fileOutputStream = new FileOutputStream(excelPath);
            wb.write(fileOutputStream);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (wb != null)
                    wb.close();
                if (fileOutputStream != null)
                    fileOutputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 创建excel，设置excel基础属性
     * @return
     * @throws Exception
     */
    private static HSSFWorkbook wbSetPrepare() throws Exception {
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.createSheet("xmlToExcel");  //声明一个sheet并命名
        sheet.setDefaultColumnWidth((short) 15);  //设置默认长度
        HSSFCellStyle style = wb.createCellStyle(); // 生成一个样式
        style.setAlignment(HorizontalAlignment.CENTER);   //样式字体居中
        toExportExcelPrepare(sheet, style, xmlPath);
        return wb;
    }

    /**
     * 做导出到excel的准备工作
     * @param sheet
     * @param style
     * @param xmlPath
     * @throws Exception
     */
    private static void toExportExcelPrepare(HSSFSheet sheet, HSSFCellStyle style, String xmlPath) throws Exception {
        HSSFRow row = sheet.createRow(0);   //创建第一行（也可以称为表头）
        HSSFCell cell = row.createCell((short) 0);//给表头第一行一次创建单元格
        xmlVo=clazz.newInstance();
        fields = clazz.getDeclaredFields();
        for (Field field :
                fields) {
            fieldString.put(field.getName(), field.getType());
        }
        for (int i = 0, j = fields.length; i < j; i++) {
            cell = row.createCell(i);
            cell.setCellValue(fields[i].getName());
            cell.setCellStyle(style);
        }
        parseXml(xmlPath);
        setExcelValue(sheet);
    }

    /**
     * 设置对应的值
     * @param sheet
     * @throws Exception
     */
    private static void setExcelValue(HSSFSheet sheet) throws Exception {
        for (int i = 0, j = xmlVoList.size(); i < j; i++) {
            HSSFRow row = sheet.createRow(i + 1);
            Object xmlVo = xmlVoList.get(i);
            for (int k = 0, l = fields.length; k < l; k++) {
                HSSFCell cell = row.createCell(k);
                String methodName = "get" + toUpperCase4Index(fields[k].getName());
                Method method = xmlVo.getClass().getDeclaredMethod(methodName);
                invokeGetMethod(method, xmlVo, fields[k].getName(), cell);
            }
        }
    }

    private static void parseXml(String xmlPath) {
        SAXReader reader = new SAXReader();
        Document document = null;
        try {
            document = reader.read(new File(xmlPath));
            Element root = document.getRootElement();
            listNodes(root);
            System.out.println();
        } catch (DocumentException e) {
            e.printStackTrace();
        }
    }

    private static void listNodes(Element element) {
        //获取节点的所有属性
        List<Attribute> attributes = element.attributes();
        for (Attribute attr :
                attributes) {
            System.out.println("节点名字为" + element.getName() + "节点属性名为" + attr.getName() + "节点属性值为" + attr.getValue());
        }
        try {
            if (fieldString.containsKey(element.getName())) {
                if (count != 0 && count % fields.length == 0) {
                    xmlVoList.add(xmlVo);
                    xmlVo = clazz.newInstance();
                }
                Method method = clazz.getDeclaredMethod("set" + toUpperCase4Index(element.getName()), fieldString.get(element.getName()));
                invokeSetMethod(xmlVo, element, method);
                count++;
            }
        } catch (Exception e) {
            System.out.println(element.getName() + "------------");
            e.printStackTrace();
        }
        Iterator elementIterator = element.elementIterator();
        while (elementIterator.hasNext()) {
            Element node = (Element) elementIterator.next();
            listNodes(node);
        }

    }

    private static void invokeGetMethod(Method method, Object xmlVo, String arg, HSSFCell cell) throws InvocationTargetException, IllegalAccessException {
        Class clazz = fieldString.get(arg);
        Object value = method.invoke(xmlVo);
        if (value != null) {
            if (clazz == byte.class || clazz == Byte.class) {
                cell.setCellValue((Byte) value);
            } else if (clazz == short.class || clazz == Short.class) {
                cell.setCellValue((Short) value);
            } else if (clazz == int.class || clazz == Integer.class) {
                cell.setCellValue((Integer) value);
            } else if (clazz == long.class || clazz == Long.class) {
                cell.setCellValue((Long) value);
            } else if (clazz == char.class || clazz == Character.class) {//用String吧，char好像没什么意义
                cell.setCellValue((Character) value);
            } else if (clazz == float.class || clazz == Float.class) {
                cell.setCellValue((Float) value);
            } else if (clazz == double.class || clazz == Double.class) {
                cell.setCellValue((Double) value);
            } else if (clazz == boolean.class || clazz == Boolean.class) {
                cell.setCellValue((Boolean) value);
            } else if (clazz == String.class) {
                cell.setCellValue(String.valueOf(value));
            }
        }
    }

    private static void invokeSetMethod(Object xmlVo, Element element, Method method) throws InvocationTargetException, IllegalAccessException {
        Class clazz = fieldString.get(element.getName());
        String value = element.getText();
        if (clazz == byte.class || clazz == Byte.class) {
            method.invoke(xmlVo, Byte.valueOf(value));
        } else if (clazz == short.class || clazz == Short.class) {
            method.invoke(xmlVo, Short.valueOf(value));
        } else if (clazz == int.class || clazz == Integer.class) {
            method.invoke(xmlVo, Integer.parseInt(value));
        } else if (clazz == long.class || clazz == Long.class) {
            method.invoke(xmlVo, Long.valueOf(value));
        } else if (clazz == char.class || clazz == Character.class) {
            method.invoke(xmlVo, value.charAt(0));
        } else if (clazz == float.class || clazz == Float.class) {
            method.invoke(xmlVo, Float.valueOf(value));
        } else if (clazz == double.class || clazz == Double.class) {
            method.invoke(xmlVo, Double.parseDouble(value));
        } else if (clazz == boolean.class || clazz == Boolean.class) {
            method.invoke(xmlVo, value.equals("true"));
        } else if (clazz == String.class) {
            method.invoke(xmlVo, value);
        }
    }


    /**
     * 首字母大写
     *
     * @param string
     * @return
     */
    private static String toUpperCase4Index(String string) {
        char[] methodName = string.toCharArray();
        methodName[0] = toUpperCase(methodName[0]);
        return String.valueOf(methodName);
    }

    /**
     * 字符转成大写
     *
     * @param chars
     * @return
     */
    private static char toUpperCase(char chars) {
        if (97 <= chars && chars <= 122) {
            chars -= 32;
        }
        return chars;
    }


}
