package com.jz.zeus.excel.test;

import cn.hutool.core.collection.ListUtil;
import cn.hutool.core.date.DateUtil;
import cn.hutool.core.date.TimeInterval;
import cn.hutool.core.lang.Console;
import com.jz.zeus.excel.ValidationInfo;
import com.jz.zeus.excel.ZeusExcel;
import com.jz.zeus.excel.test.data.DemoData;

import java.util.*;

public class CascadeWriteTest {

  private final static String path = "/Users/jd/Desktop/254.xlsx";

  public static void main(String[] args) {
    Console.log("写入Excel前内存：{}M", (Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory())/(1024*1024));
    TimeInterval timer = DateUtil.timer();
    ZeusExcel.write(path)
          .sheet("及联")
          .validationInfos(getValidationInfo())
          .doWrite(DemoData.class, getDataList("测试_", 10));

    Console.log("写入耗时：{}s", timer.intervalSecond());
    Console.log("写入Excel后内存：{}M", (Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory())/(1024*1024));
    Console.log("end");
  }


  private static List<ValidationInfo> getValidationInfo() {
    ValidationInfo provinces = ValidationInfo.buildColumnByField(DemoData::getProvinces, "上海市", "河南省", "北京市")
          .asDicSheet("省").setDicTitle("中国的省");

    Map<String, List<String>> cityMap = new HashMap<>();
    cityMap.put("上海市", ListUtil.toList("上海市"));
    cityMap.put("河南省", ListUtil.toList("郑州市", "南阳市", "信阳市"));
    cityMap.put("北京市", ListUtil.toList("北京市"));
    ValidationInfo city = ValidationInfo.buildCascadeByField(DemoData::getCity, provinces, cityMap).asDicSheet("市");

    Map<String, List<String>> townMap = new HashMap<>();
    townMap.put("上海市", ListUtil.toList("静安区", "黄浦区", "徐汇区"));
    townMap.put("北京市", ListUtil.toList("通州区", "朝阳区", "顺义区"));
    townMap.put("河南省", ListUtil.toList("二七区", "新郑市"));
    ValidationInfo town = ValidationInfo.buildCascadeByField(DemoData::getTown, provinces, townMap).asDicSheet("区");
    return ListUtil.toList(
//            ValidationInfo.buildColumnByField("id", list).setErrorBox("Error", "请选择正确的ID"),
//            ValidationInfo.buildColumnByHead("destPlus（选填）", "是", "否"),
//            ValidationInfo.buildColumnByHead("destPlus（选填）", "是自定义", "不是自定义").asDicSheet("字典表", "说明f辅导费")
          ValidationInfo.buildColumnByField(DemoData::getSrc, "是自定义", "不是自定义"),
          provinces, city
          , town
    );
  }

  private static List<DemoData> getDataList(String prefix, int num) {
    List<DemoData> dataList = new ArrayList<>();
    for (int i = 0; i < num; i++) {
      DemoData demoData = new DemoData();
      demoData.setId(Long.valueOf(i));
      demoData.setDest(prefix + "dest" + i);
      demoData.setSrc(prefix + "src" + i);
      demoData.setFunc(prefix + "func" + i);
//            demoData.setPrice(BigDecimal.valueOf(3.94));
      demoData.setPrice(null);

      Map<String, String> map = new LinkedHashMap<>();
//            map.put("图片（必填）\n允许的文件类型*JPG,1080*1920，大小限制150K。素材必须满足腾讯所有规格要求，否则无法通过审核。", prefix + "12");
//            map.put("落地页（必填）", prefix + "jfak");
//            map.put("第三方异步点击监测URL（2）", prefix + "jfak");
//            map.put("自定义3", prefix + "集分宝");
      demoData.setExtendColumnMap(map);
      dataList.add(demoData);
    }
    return dataList;
  }

}
