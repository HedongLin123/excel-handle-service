package com.itdl.excel.controller;

import cn.hutool.http.server.HttpServerResponse;
import com.itdl.excel.core.ExcelParse;
import com.itdl.excel.util.FileZipUtil;
import com.itdl.excel.vo.ExcelDataResult;
import org.apache.commons.lang3.StringUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * @Description
 * @Author donglin.he
 * @Date 2023/04/20 11:39
 */
@RestController
public class ExcelController {
    @Value("${excel.savePath:/data/generator/excel}")
    private String excelSavePath;

    @PostMapping("/parseExcel")
    public void getExcelContent(@RequestParam("file") MultipartFile file, HttpServletResponse response) throws IOException {
        if (file.isEmpty()){
            response.reset();
            response.setCharacterEncoding("UTF-8");
            ServletOutputStream outputStream = response.getOutputStream();
            outputStream.write("upload file cannot empty".getBytes(StandardCharsets.UTF_8));
            outputStream.flush();
            return;
        }

        InputStream is = file.getInputStream();
        List<ExcelDataResult> alls = ExcelParse.simpleRead(is);

        AtomicInteger countTotal = new AtomicInteger(1);

        for (ExcelDataResult all : alls) {
            AtomicInteger count = new AtomicInteger(1);
            Map<Integer, String> tableHead = new HashMap<>();
            tableHead = all.getDatas().get(0);

            Map<Integer, String> map = null;

            // 从哪一行进行分割，因为是要在lambda读取splitHeadNum原子操作
            Queue<Integer> queue = new LinkedList<>();

            // 将Excel进行分割（企业信息与内容进行分割）
            for (int i = 0; i < all.getDatas().size(); i++) {
                int temp = i;
                map = all.getDatas().get(i);
                map.forEach((k, v) -> {
                    if (v != null) {
                        if ("小计".equals(v)) {
                            // 获取头部要分割的行数
                            queue.add(temp);
                        }
                    }
                });
            }
            List<List<Map<Integer, String>>> subListMap = new ArrayList<>();
            Integer curr = 1;
            Integer next = 1;
            while (!queue.isEmpty()){
                final Integer poll = queue.poll();
                curr = next ;
                next = poll;
                List<Map<Integer, String>> subMap = new ArrayList<>();
                for (int i = curr; i <= next; i++) {
                    final Map<Integer, String> map1 = all.getDatas().get(i);
                    System.out.println(map1);
                    subMap.add(map1);
                }
                next++;
                System.out.println("========================================================================================");
                subListMap.add(subMap);
                final int currCount = count.incrementAndGet();
                final int total = countTotal.get();
                final Map head = all.getHead();

                // handle content is double
                subMap = handleSubMap(subMap);

                File file1 = new File(excelSavePath);
                if (!file1.exists()){
                    file1.mkdirs();
                }

                ExcelParse.generatorExcel(all.getSheetName(), head, tableHead, subMap, excelSavePath + "/subExcel-" + total + "-" + currCount + ".xlsx");
            }

//            System.out.println(subListMap);
            countTotal.incrementAndGet();
        }

        FileZipUtil.exportZip(response, excelSavePath, "过磅明细管理", ".zip");
        File file1 = new File(excelSavePath);
        File[] files = file1.listFiles();
        for (File file2 : files) {
            file2.delete();
        }
    }

    private List<Map<Integer, String>> handleSubMap(List<Map<Integer, String>> subMap) {
        List<Map<Integer, String>> resultMap = new ArrayList<>();
        for (Map<Integer, String> map : subMap) {
            Map<Integer, String> newMap = new HashMap<>();
            for (Integer key : map.keySet()) {
                final String value = map.get(key);
                try {
                    final BigDecimal bigDecimal = new BigDecimal(value).setScale(2, BigDecimal.ROUND_HALF_UP);
                    String newValue = String.valueOf(bigDecimal);
                    if (newValue.endsWith(".00")){
                        // 整数，不要小数点
                        newValue = String.valueOf(Long.valueOf(newValue));
                    }
                    newMap.put(key, newValue);
                }catch (Exception e){
                    newMap.put(key, value);
                }
            }
            resultMap.add(newMap);
        }

        return resultMap;
    }

}
