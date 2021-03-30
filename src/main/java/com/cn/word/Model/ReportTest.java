package com.cn.word.Model;

import com.deepoove.poi.data.ChartMultiSeriesRenderData;
import com.deepoove.poi.data.ChartSingleSeriesRenderData;
import lombok.Data;

import java.util.List;


@Data
public class ReportTest {

    private List<Label> label;//列表
    private ChartMultiSeriesRenderData pieChartOne;//条形图
    private ChartSingleSeriesRenderData pieChartTwo;//饼图



    @Data
    public static class Label{
        private Integer number;
        private String name;
        private Integer age;
        private String address;
        private String time;
    }

    public String[] getHackLoopTableFiles(){
        String[] names=new String[]{
                "label"
        };
        return names;
    }

    public String[] getDecimalFormatFiles(){
        String[] names=new String[]{
                "number",
                "name",
                "age",
                "address",
                "time"
        };
        return names;
    }
}


