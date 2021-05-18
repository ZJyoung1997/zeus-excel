package com.jz.zeus.excel.test;

import cn.hutool.core.collection.ListUtil;
import cn.hutool.core.lang.Console;
import com.jz.zeus.excel.test.data.A;
import com.jz.zeus.excel.test.data.B;
import com.jz.zeus.excel.test.data.C;
import com.jz.zeus.excel.test.data.DemoData;
import com.jz.zeus.excel.util.ValidatorUtils;
import com.jz.zeus.excel.validator.BeanValidator;
import com.jz.zeus.excel.validator.VerifyResult;
import lombok.SneakyThrows;

/**
 * @author:JZ
 * @date:2021/3/24
 */
public class ValidatorTest {

    public static void main(String[] args) {
//        test1();
        test2();
    }

    @SneakyThrows
    public static void test2() {
        BeanValidator<B> beanValidator = BeanValidator.<B>build()
                .nonNull(B::getId, "B id 不能为空")
                .verify(B::getCity, e -> "名字".equals(e), "B city 必须为 名字");
        BeanValidator cValidator = BeanValidator.<C>build()
                .nonNull(C::getId, "id 不能为空")
                .verify(C::getTown, e -> "城镇".equals(e), "c 的 town 只能为 城镇");
        BeanValidator bValidator = BeanValidator.<B>build()
                .nonNull(B::getId, "B 的id不能为空")
                .verify(B::getCity, e -> "城市".equals(e), "B 的 city 只能为 城市")
                .verifyCollection(B::getCList, cValidator);

        B b1 = new B();
        b1.setId(1L);
        b1.setCity("城市");
        b1.setCList(ListUtil.toList(new C(99L, "城镇"), new C(100L, "非城镇")));
        B b2 = new B();
        b2.setCity("城市1");
        b2.setCList(ListUtil.toList(new C(99L, "1城镇"), new C(100L, "非城镇")));

        A a = new A();
        a.setId(null);
        a.setName("jkj");
        a.setB(b2);
        a.setBList(ListUtil.toList(b1, b2));
//        VerifyResult verifyResult = BeanValidator.build(a)
//                .nonNull(A::getId, "A 的id不能为空")
//                .verifyCollection(A::getBList, bValidator)
//                .doVerify();

        VerifyResult verifyResult1 = BeanValidator.build(a)
                .verifyBean(A::getB, beanValidator)
                .verifyByAnnotation(A::getB)
                .doVerify(true);
        Console.log(verifyResult1);
    }

    public static void test1() {
        DemoData demoData = new DemoData();
        demoData.setSrc("02645");
        System.out.println(ValidatorUtils.validate(demoData));
    }

}
