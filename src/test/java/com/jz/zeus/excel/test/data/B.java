package com.jz.zeus.excel.test.data;

import lombok.Data;
import org.hibernate.validator.constraints.Length;

import java.util.List;

/**
 * @Author JZ
 * @Date 2021/5/18 13:22
 */
@Data
public class B {

    private Long id;

    @Length(min = 1, max = 2, message = "city 最大长度不能超过2")
    private String city;

    private List<C> cList;

}
