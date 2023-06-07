package mao;

import cn.afterturn.easypoi.word.WordExportUtil;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.FileOutputStream;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Project name(项目名称)：java报表_EasyPOI导出word
 * Package(包名): mao
 * Class(类名): Test2
 * Author(作者）: mao
 * Author QQ：1296193245
 * GitHub：https://github.com/maomao124/
 * Date(创建日期)： 2023/6/7
 * Time(创建时间)： 16:14
 * Version(版本): 1.0
 * Description(描述)： 无
 */

public class Test2
{
    /**
     * 得到int随机
     *
     * @param min 最小值
     * @param max 最大值
     * @return int
     */
    public static int getIntRandom(int min, int max)
    {
        if (min > max)
        {
            min = max;
        }
        return min + (int) (Math.random() * (max - min + 1));
    }

    public static void main(String[] args) throws Exception
    {
        //数据
        Map<String, Object> params = new HashMap<>();
        params.put("name", "员工信息表");

        //下面是表格中需要的数据
        List<Map<String, Object>> mapList = new ArrayList<>();
        Map<String, Object> map = null;
        for (int i = 1; i <= 180; i++)
        {
            map = new HashMap<>();
            map.put("id", i);
            map.put("name", "姓名" + i);
            map.put("age", getIntRandom(15, 30));
            map.put("address", "中国");
            mapList.add(map);
        }
        //把组建好的表格需要的数据放到大map中
        params.put("mapList", mapList);

        params.put("date_year", LocalDate.now().getYear());
        params.put("date_month", LocalDate.now().getMonthValue());
        params.put("date_day", LocalDate.now().getDayOfMonth());

        //写入
        XWPFDocument xwpfDocument = WordExportUtil.exportWord07("./template2.docx", params);
        FileOutputStream fileOutputStream = new FileOutputStream("out2.docx");
        xwpfDocument.write(fileOutputStream);
        xwpfDocument.close();
        xwpfDocument.close();

    }
}
