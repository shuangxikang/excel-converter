package com.ksx.tools.excel;

import com.ksx.tools.excel.annotation.ExcelColumn;

import java.math.BigDecimal;
import java.util.Date;

/**
 * Excel 导入对象
 * Created by ksx on 2019-01-28.
 */
public class ExcelImportVO {
    /* 用户账号 */
    @ExcelColumn(name = "用户账号", column = "A")   //注：name=Excel导出时表头名称、column=对应Excel列编号一一对应
    private String userAccount;
    /* 年龄 */
    @ExcelColumn(name = "年龄", column = "B")
    private int age;
    /* 生日 */
    @ExcelColumn(name = "生日", column = "D")
    private Date birthday;
    /* 金额*/
    @ExcelColumn(name = "金额：￥", column = "F", dataPattern = "0.000")   //注：dataFormat=数字格式化模式字符串
    private BigDecimal amount;

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

    public Date getBirthday() {
        return birthday;
    }

    public void setBirthday(Date birthday) {
        this.birthday = birthday;
    }

    public BigDecimal getAmount() {
        return amount;
    }

    public void setAmount(BigDecimal amount) {
        this.amount = amount;
    }
}
