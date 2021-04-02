package com.carbonfive.sstemplates.tags;

import com.carbonfive.sstemplates.SsTemplateContext;
import com.carbonfive.sstemplates.SsTemplateException;


public interface SsTemplateTag {

    void render (SsTemplateContext context) throws SsTemplateException;

    void addTag (SsTemplateTag tag);
}
