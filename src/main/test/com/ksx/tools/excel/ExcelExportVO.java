package com.ksx.tools.excel;

import com.ksx.tools.excel.annotation.ExcelColumn;
import com.ksx.tools.excel.annotation.ExcelColumnSplitCell;
import com.ksx.tools.excel.format.DataFormat;
import com.ksx.tools.excel.format.number.NumberToStringFormat;
import com.ksx.tools.excel.style.RedDataCellStyle;
import com.ksx.tools.excel.style.RedHeaderCellStyle;
import com.ksx.tools.excel.style.YellowHeaderCellStyle;

import java.math.BigDecimal;
import java.util.Date;
import java.util.List;

/**
 * Excel 导入对象
 * Created by kangshuangxi on 2017/1/17.
 */
public class ExcelExportVO {
    /* 用户账号 */
    @ExcelColumn(name = "用户账号", column = "A")   //注：name=Excel导出时表头名称、column=对应Excel列编号一一对应
    private String userAccount;
    /* 年龄 */
    @ExcelColumn(name = "年龄", column = "B", headerCellStyle = RedHeaderCellStyle.class, dataCellStyle = RedDataCellStyle.class) //注：用headerCellStyle和dataCellStyle属性分别设置表头和数据单元格样式，实现对应接口协议即可自定义
    private int age;
    /* 性别：F=男、M=女 */
    @ExcelColumn(name = "性别", column = "C", prompt = "F=男\nM=女", dataFormat = GenderDataFormat.class)   //注：prompt=批注信息，用户鼠标点上去会给出tip信息
    private char gender;
    /* 生日 */
    @ExcelColumn(name = "生日", column = "D", dataPattern = "yyyy/MM/dd") //注：datePattern=日期显示模式字符串
    private Date birthday;
    /* 密码 */
    @ExcelColumn(name = "密码", column = "E", isExport = false, headerCellStyle = YellowHeaderCellStyle.class)   //注：isExport=是否导出本列数据
    private String password;
    /* 金额*/
    @ExcelColumn(name = "金额：￥", column = "F", dataPattern = "0.000")   //注：dataFormat=数字格式化模式字符串
    private BigDecimal amount;
    /* 身份证 */
    @ExcelColumn(name = "身份证号", column = "G", dataFormat = NumberToStringFormat.class)   //注：数字转为字符串；dataFormat=特殊类型的数据格式化类，实现DataFormat<T>接口即可
    private long idCard;
    /* 单元格纵向拆分，具体拆分行数等于列表长度 */
    @ExcelColumn(column = "H")
    @ExcelColumnSplitCell( beginColumn = "H", endColumn = "J", itemClass = Address.class)
    private List<Address> addressList;
    /* 状态标示：1、未激活；2、正常；3、存在异常；4、已锁定（禁用）；5、删除 */
    @ExcelColumn(name = "状态标示", column = "K", dataFormat = StatusDataFormat.class)  //dataFormat=这里用数据格式化把状态转换为可理解字符串，可以用来转换任何不易理解的定义字段
    private int status;

    public String getUserAccount() {
        return userAccount;
    }

    public void setUserAccount(String userAccount) {
        this.userAccount = userAccount;
    }

    public int getAge() {
        return age;
    }

    public void setAge(int age) {
        this.age = age;
    }

    public char getGender() {
        return gender;
    }

    public void setGender(char gender) {
        this.gender = gender;
    }

    public Date getBirthday() {
        return birthday;
    }

    public void setBirthday(Date birthday) {
        this.birthday = birthday;
    }

    public String getPassword() {
        return password;
    }

    public void setPassword(String password) {
        this.password = password;
    }

    public BigDecimal getAmount() {
        return amount;
    }

    public void setAmount(BigDecimal amount) {
        this.amount = amount;
    }

    public long getIdCard() {
        return idCard;
    }

    public void setIdCard(long idCard) {
        this.idCard = idCard;
    }

    public List<Address> getAddressList() {
        return addressList;
    }

    public void setAddressList(List<Address> addressList) {
        this.addressList = addressList;
    }

    public int getStatus() {
        return status;
    }

    public void setStatus(int status) {
        this.status = status;
    }

    /**
     * 单元格拆分（收货地址信息）
     */
    public static class Address {
        /* 姓名 */
        @ExcelColumn(column = "H", name = "姓名")
        private String name;
        /* 手机号 */
        @ExcelColumn(column = "I", name = "手机号")
        private String  mobileNumber;
        /* 地址 */
        @ExcelColumn(column = "J", name = "地址")
        private String  location;

        public Address(String name, String mobileNumber, String  location) {
            this.name = name;
            this.mobileNumber = mobileNumber;
            this.location = location;
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public String getMobileNumber() {
            return mobileNumber;
        }

        public void setMobileNumber(String mobileNumber) {
            this.mobileNumber = mobileNumber;
        }

        public String getLocation() {
            return location;
        }

        public void setLocation(String location) {
            this.location = location;
        }
    }

    /**
     * 导出数据自定义格式化（这里把性别字符"'M'、'F'"简单转换为：男、女、保密）
     * @param <T>
     */
    public static class GenderDataFormat<T> implements DataFormat<T> {
        @Override
        public String format(T data) {
            if (((Character) data).charValue() == 'M') {
                return "男";
            } else if (((Character) data).charValue() == 'F') {
                return "女";
            } else {
                return "保密";
            }
        }
    }

    /**
     * 导出数据自定义格式化（把数字导出为容易理解等字符串）
     * 这里为了使demo简介状态直接用魔数定义了，实际操作中尽量使用工具类内部封装为枚举定义
     * @param <T>
     */
    public static class StatusDataFormat<T> implements DataFormat<T> {

        @Override
        public String format(T data) {
            int status = (Integer) data;
            switch (status) {
                case 1:
                    return "未激活";
                case 2:
                    return "正常";
                case 3:
                    return "存在异常";
                case 4:
                    return "已锁定（禁用）";
                case 5:
                    return "删除";
            }

            return "";
        }
    }
}
