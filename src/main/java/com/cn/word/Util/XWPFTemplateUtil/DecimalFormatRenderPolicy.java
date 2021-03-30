package com.cn.word.Util.XWPFTemplateUtil;

import com.deepoove.poi.policy.AbstractRenderPolicy;
import com.deepoove.poi.render.RenderContext;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.text.DecimalFormat;


public class DecimalFormatRenderPolicy extends AbstractRenderPolicy<Object> {
    private String pattern="#,###";
    public DecimalFormatRenderPolicy(){

    }
    public DecimalFormatRenderPolicy(String pattern){
        this.pattern=pattern;
    }
    @Override
    public void doRender(RenderContext<Object> context) throws Exception {
        XWPFRun where = context.getWhere();
        // anything
        Object thing = context.getThing();
        // do 文本
//        DecimalFormat df = new DecimalFormat(pattern);

//        where.setText(df.format(thing), 0);
        where.setText(thing.toString(), 0);
    }

}
