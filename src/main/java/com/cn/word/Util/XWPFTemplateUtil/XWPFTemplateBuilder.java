package com.cn.word.Util.XWPFTemplateUtil;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.config.ConfigureBuilder;
import com.deepoove.poi.policy.HackLoopTableRenderPolicy;
import com.deepoove.poi.policy.RenderPolicy;


public class XWPFTemplateBuilder {
    private XWPFTemplate template;
    private ConfigureBuilder configure;
    private String fileName;

    private XWPFTemplateBuilder(){

    }
    public XWPFTemplateBuilder(String fileName){
        setFileName(fileName);
        initConfigure();
    }

    private void initConfigure(){
        configure = Configure.newBuilder();
    }

    public XWPFTemplateBuilder setFileName(String fileName){
        this.fileName=fileName;
        return this;
    }

    public XWPFTemplateBuilder bind(String tagName, RenderPolicy policy){
        configure.bind(tagName,policy);
        return this;
    }

    public XWPFTemplateBuilder bindHackLoopTableRenderPolicy(String tagName){
        configure.bind(tagName,new HackLoopTableRenderPolicy());
        return this;
    }

    public XWPFTemplateBuilder bindHackLoopTableRenderPolicy(String[] tagNames){
        if(null==tagNames) {return this;}
        HackLoopTableRenderPolicy policy= new HackLoopTableRenderPolicy();
        for (int i=0;i<tagNames.length;i++) {
            configure.bind(tagNames[i], policy);
        }
        return this;
    }

    public XWPFTemplateBuilder bindDecimalFormatRenderPolicy(String tagName){
        configure.bind(tagName,new DecimalFormatRenderPolicy());
        return this;
    }

    public XWPFTemplateBuilder bindDecimalFormatRenderPolicy(String[] tagNames){
        if(null==tagNames) {return this;}
        DecimalFormatRenderPolicy policy = new DecimalFormatRenderPolicy();
        for (int i=0;i<tagNames.length;i++) {
            configure.bind(tagNames[i], policy);
        }
        return this;
    }
    public XWPFTemplate builder(){
        fileName=this.getClass().getResource("/template/"+fileName).getPath();
        template = XWPFTemplate.compile(fileName,configure.build());
        return template;
    }
}
