package com.xxkj.demo.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

import com.xxkj.demo.entity.WordGainPrizeDto;
import freemarker.log.Logger;
import freemarker.template.Configuration;
import freemarker.template.TemplateExceptionHandler;
import sun.misc.BASE64Encoder;

import javax.servlet.http.HttpServletResponse;

/**
 * @Description 利用FreeMarker导出Word
 */
public class WordUtil {
    private Logger log = Logger.getLogger(WordUtil.class.toString());

    public String createDoc(Map dataMap, String path, String srcFileName) {
        String fileName = UUID.randomUUID().toString() + ".doc";
        HttpServletResponse httpServletResponse = ServletCommonUtil.getHttpServletResponse();
        try {
            httpServletResponse.setHeader("Content-Disposition",
                    "attachment; filename=\"" + new String(fileName.getBytes("gbk"), "iso-8859-1") + "\"");
            httpServletResponse.setContentType("application/octet-stream;charset=UTF-8");

            // 设置模本装置方法和路径,FreeMarker支持多种模板装载方法。可以重servlet，classpath，数据库装载，
            /*Configuration configuration = new Configuration();
            configuration.setDefaultEncoding("utf-8");
            configuration.setDirectoryForTemplateLoading(new File(filePath));
            Template t = configuration.getTemplate("证明.ftl");
            Writer writer = httpServletResponse.getWriter();
            t.process(dataMap, new BufferedWriter(writer));*/

            FreemarkerUtils.build(this.getClass(), path).setTemplate(srcFileName).generate(dataMap,
                    httpServletResponse.getOutputStream());

        } catch (Exception e) {
            e.printStackTrace();
        }

        return fileName;
    }

    public static String getRootClassPath() {
        String rootClassPath = "";
        try {
            String path = WordUtil.class.getClassLoader().getResource("").toURI().getPath();
            rootClassPath = new File(path).getAbsolutePath();
        } catch (Exception e) {
            String path = WordUtil.class.getClassLoader().getResource("").getPath();
            rootClassPath = new File(path).getAbsolutePath();
        }
        return rootClassPath;
    }

    /**
     * 获得图片的Base64编码
     *
     * @param imgFile
     * @return
     */
    public String getImageStr(String imgFile) {
        //imgFile = getRootClassPath() + imgFile;
        imgFile = "D:/xdkj/git-code/abtc/new/abtc_scholarship-back-v1.0/target/classes/file/test.jpg";
        InputStream in = null;
        byte[] data = null;
        try {
            in = new FileInputStream(imgFile);
        } catch (FileNotFoundException e) {
            log.error("加载图片未找到", e);
            e.printStackTrace();
        }
        try {
            data = new byte[in.available()];
            //注：FileInputStream.available()方法可以从输入流中阻断由下一个方法调用这个输入流中读取的剩余字节数
            in.read(data);
            in.close();
        } catch (IOException e) {
            log.error("IO操作图片错误", e);
            e.printStackTrace();
        }
        BASE64Encoder encoder = new BASE64Encoder();
        return encoder.encode(data);

    }

    public static void main(String[] args) {
        WordUtil emw = new WordUtil();
        Map dataMap = new HashMap();
        // 基本信息
        dataMap.put("image", new WordUtil().getImageStr("/file/test.jpeg"));

        dataMap.put("college", "test");
        dataMap.put("class", "1版");
        dataMap.put("sno", "123456");
        dataMap.put("name", "黄xx");
        dataMap.put("age", 26);
        dataMap.put("yearMonth", "1999/12");
        dataMap.put("political", "群众");
        dataMap.put("nation", "汉");
        dataMap.put("onScTime", "1997/09/01");
        dataMap.put("major", "111");
        dataMap.put("phone", "222");
        dataMap.put("idCard", "333");
        dataMap.put("birthPlace", "444");
        dataMap.put("homeAddress", "555");
        // 主要获奖情况
        List list = new ArrayList();
        list.add(new WordGainPrizeDto() {{
            setD("wwwww");
            setPrizeName("ssssss");
            setAwardUnit("xxxxxx");
        }});
        dataMap.put("list", list);


        /*dataMap.put("d1", "666");
        dataMap.put("prizeName1", "777");
        dataMap.put("awardUnit1", "888");

        dataMap.put("d2", "666");
        dataMap.put("prizeName2", "777");
        dataMap.put("awardUnit2", "888");

        dataMap.put("d3", "666");
        dataMap.put("prizeName3", "777");
        dataMap.put("awardUnit3", "888");

        dataMap.put("d4", "666");
        dataMap.put("prizeName4", "777");
        dataMap.put("awardUnit4", "888");

        dataMap.put("d5", "666");
        dataMap.put("prizeName5", "777");
        dataMap.put("awardUnit5", "888");

        dataMap.put("d6", "666");
        dataMap.put("prizeName6", "777");
        dataMap.put("awardUnit6", "888");

        dataMap.put("d7", "666");
        dataMap.put("prizeName7", "777");
        dataMap.put("awardUnit7", "888");

        dataMap.put("d8", "666");
        dataMap.put("prizeName8", "777");
        dataMap.put("awardUnit8", "888");*/


        // 在校期间主要表现
        dataMap.put("performance", "999");
        dataMap.put("pyear", "2022");
        dataMap.put("pmonth", "02");
        dataMap.put("pday", "01");
        // 年级推荐意见
        dataMap.put("gradeRecom", "11111");
        dataMap.put("gyear", "2022");
        dataMap.put("gmonth", "07");
        dataMap.put("gday", "02");
        // 学院审查意见
        dataMap.put("collegeCheck", "22222");
        dataMap.put("cyear", "2022");
        dataMap.put("cmonth", "09");
        dataMap.put("cday", "02");
        // 学校审批意见
        dataMap.put("publicDay", "33333");
        dataMap.put("schoolCheck", "44444");
        dataMap.put("syear", "2022");
        dataMap.put("smonth", "09");
        dataMap.put("sday", "08");
        //this.getClass().getResource("/");
        emw.createDoc(dataMap, "/file/", "recommendationForm.ftl");
    }
}