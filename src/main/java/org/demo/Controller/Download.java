package org.demo.Controller;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.demo.Class.Sample;
import org.demo.Util.ExcelBuilderUtil;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

@Controller
public class Download {
    
    @RequestMapping("/download.xlsx")
    public ModelAndView download() {
        
        List<String> header1 = createHeader();
        List<String> header2 = createHeader();
        Map<Integer, List<String>> header = new HashMap<>();
        header.put(0, header1);
        header.put(1, header2);
        List<String> sheet = createSheet();
        List<Sample> sample1 = createContent();
        List<Sample> sample2 = createContent();
        Map<Integer, Object> content = new HashMap<>();
        content.put(0, sample1);
        content.put(1, sample2);
        
        Map<String, Object> map = new HashMap<>();
        map.put("headers", header);
        map.put("contents", content);
        map.put("sheets", sheet);
        
        return new ModelAndView(new ExcelBuilderUtil(), map);
    }
    
    private List<String> createHeader() {
        
        List<String> headers = new ArrayList<>();
        headers.add("ID");
        headers.add("NAME");
        
        return headers;
    }
    
    private List<String> createSheet() {
        
        List<String> sheets = new ArrayList<>();
        sheets.add("sheet1");
        sheets.add("sheet2");
        
        return sheets;
    }
    
    private List<Sample> createContent() {
        
        List<Sample> contents = new ArrayList<>();
        for(int i = 0; i < 5; i++) {
            Sample sample = new Sample();
            sample.setId(i);
            sample.setName("sample_" + i);
            contents.add(sample);
        }
        
        return contents;
    }
}