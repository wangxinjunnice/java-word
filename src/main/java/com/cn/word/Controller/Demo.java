package com.cn.word.Controller;

import com.cn.word.Model.ReportTest;
import com.cn.word.Util.XWPFTemplateUtil.XWPFTemplateBuilder;
import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.data.ChartMultiSeriesRenderData;
import com.deepoove.poi.data.ChartSingleSeriesRenderData;
import com.deepoove.poi.data.SeriesRenderData;
import com.deepoove.poi.policy.HackLoopTableRenderPolicy;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.net.URLEncoder;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;


@RestController
@RequestMapping("/test")
public class Demo {


    @GetMapping(value = "/report")
    public  void testOne(HttpServletResponse response) {

        ReportTest reportTest=new ReportTest();
        DateTimeFormatter fmt = DateTimeFormatter.ofPattern("yyyy年MM月dd日");
        String format = LocalDate.now().format(fmt);

        List<String> categories=null;

                //条形图
        categories=new ArrayList<>();
        categories.add("北京");
        categories.add("陕西");
        categories.add("上海");

        List<SeriesRenderData> list=new ArrayList<>();
        SeriesRenderData seriesRenderData = new SeriesRenderData("数量",new Integer[]{10,3,5});
        list.add(seriesRenderData);

        ChartMultiSeriesRenderData tiaoPieChart = new ChartMultiSeriesRenderData ();
        tiaoPieChart.setChartTitle("985大学排名");
        tiaoPieChart.setCategories(categories.toArray(new String[categories.size()]));
        tiaoPieChart.setSeriesDatas(list);
        reportTest.setPieChartOne(tiaoPieChart);

        //列表
        List<String> listName = Arrays.asList("刘德华","黎明","郭富城","张学友");
        List<String> listAddress = Arrays.asList("北京","上海","深圳","广东");
        List<Integer> ags = Arrays.asList(50,52,57,60);
        List<ReportTest.Label> list1=new ArrayList<>();
        for (int i = 0; i < 4; i++) {
            ReportTest.Label label=new ReportTest.Label();
            label.setNumber(i+1);
            label.setName(listName.get(i));
            label.setAddress(listAddress.get(i));
            label.setAge(ags.get(i));
            label.setTime(format);
            list1.add(label);
        }
        reportTest.setLabel(list1);

        //饼图
        categories=new ArrayList<>();
        categories.add("北京");
        categories.add("上海");
        categories.add("深圳");
        categories.add("广东");
        ChartSingleSeriesRenderData pieChart = new ChartSingleSeriesRenderData();
        pieChart.setChartTitle("一线城市");
        pieChart.setCategories(categories.toArray(new String[categories.size()]));
        pieChart.setSeriesData(new SeriesRenderData("数量",new Integer[]{25,25,25,25}));
        reportTest.setPieChartTwo(pieChart);

        XWPFTemplateBuilder template=new XWPFTemplateBuilder("test.docx");
        template.bindHackLoopTableRenderPolicy(reportTest.getHackLoopTableFiles());
        template.bindDecimalFormatRenderPolicy(reportTest.getDecimalFormatFiles());


        XWPFTemplate render = template.builder().render(reportTest);
        try {
            //Response回写文件
            response.reset();
            response.setContentType("application/octet-stream; charset=utf-8");
            response.setHeader("Content-Disposition", "attachment; filename=" + URLEncoder.encode("测试"+".docx", "UTF-8"));

            render.write(response.getOutputStream());
            response.flushBuffer();
            render.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
