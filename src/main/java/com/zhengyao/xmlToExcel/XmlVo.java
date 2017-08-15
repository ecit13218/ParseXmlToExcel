package com.zhengyao.xmlToExcel;

/**
 * Created by 郑瑶 on 2017/8/12.
 */
public class XmlVo {
    private int maxLength;
    private String name;
    private String type;
    private Character charac;
    private boolean booleans;

    public boolean getBooleans() {
        return booleans;
    }

    public void setBooleans(boolean booleans) {
        this.booleans = booleans;
    }

    public Character getCharac() {
        return charac;
    }

    public void setCharac(Character charac) {
        this.charac = charac;
    }
//    public String getField() {
//        return field;
//    }
//
//    public void setField(String field) {
//        this.field = field;
//    }


    public int getMaxLength() {
        return maxLength;
    }

    public void setMaxLength(int maxLength) {
        this.maxLength = maxLength;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }
}
