package com.leyunon.demoeasyexcel;


import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.enums.WriteDirectionEnum;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.fill.FillConfig;
import com.alibaba.excel.write.metadata.fill.FillWrapper;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class DemoEasyExcel1 {

    public static void main(String[] args) throws IOException {
        DemoEasyExcel1 modelExportBug = new DemoEasyExcel1();
        modelExportBug.modelBug();
    }

    public void modelBug() throws IOException {
        File file = new File("export-mode.xlsx");
        File tempFile = new File("f://test.xlsx");
        ExcelWriter excelWriter = EasyExcel.write(tempFile)
                .withTemplate(new FileInputStream(file))
                .build();
        WriteSheet writeSheet = EasyExcel.writerSheet().build();
        FillConfig horizontalFillConfig = FillConfig.builder().direction(WriteDirectionEnum.HORIZONTAL).build();
        FillConfig verticalFillConfig = FillConfig.builder().forceNewRow(false).direction(WriteDirectionEnum.VERTICAL).build();

        TestHeadData testHeadData = new TestHeadData();
        testHeadData.setNoCount("头部信息");
        testHeadData.setCreateTime("头部信息");
        testHeadData.setRemark("头部信息");
        testHeadData.setSumCount("头部信息");
        ArrayList<TestHeadData> head = new ArrayList<>();
        head.add(testHeadData);
        excelWriter.fill(new FillWrapper(TestHeadData.class.getSimpleName(), head), horizontalFillConfig, writeSheet);

        List<TestData> testDataList = new ArrayList<>();
        for (int i = 1; i <= 10; i++) {
            TestData testData = new TestData();
            testData.setColor("颜色");
            testData.setContent("内容");
            testData.setCount("数目");
            testData.setKey("key" + i);
            testData.setName("name" + i);
            testData.setNo(i + "");
            testData.setRemark(" ");
            testData.setTestInfo("信息" + i);
            testData.setType(i + "");
            testDataList.add(testData);
        }

        excelWriter.fill(new FillWrapper(TestData.class.getSimpleName(), testDataList), verticalFillConfig, writeSheet);
        excelWriter.finish();
    }


    /**
     * 导出
     *
     * @param
     * @param templatePath
     */
    public static ExcelWriter buildExport(File file,
                                          InputStream templatePath) throws IOException {
        return EasyExcel.write(file)
                .withTemplate(templatePath)
                .build();
    }
}
