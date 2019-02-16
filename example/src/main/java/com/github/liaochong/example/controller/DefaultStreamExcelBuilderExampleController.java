package com.github.liaochong.example.controller;

import com.github.liaochong.example.pojo.ArtCrowd;
import com.github.liaochong.html2excel.core.DefaultStreamExcelBuilder;
import org.apache.commons.codec.CharEncoding;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.net.URLEncoder;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.Executors;

/**
 * @author liaochong
 * @version 1.0
 */
@Controller
public class DefaultStreamExcelBuilderExampleController {

    @GetMapping("/default/excel/stream/example")
    public void streamBuild(HttpServletResponse response) throws Exception {
        DefaultStreamExcelBuilder defaultExcelBuilder = DefaultStreamExcelBuilder.of(ArtCrowd.class)
                .threadPool(Executors.newFixedThreadPool(10))
                .start();

        List<CompletableFuture> futures = new ArrayList<>();
        for (int i = 0; i < 100; i++) {
            CompletableFuture future = CompletableFuture.runAsync(() -> {
                List<ArtCrowd> dataList = this.getDataList();
                defaultExcelBuilder.append(dataList);
            });
            futures.add(future);
        }
        futures.forEach(CompletableFuture::join);
        Workbook workbook = defaultExcelBuilder.build();

        response.setCharacterEncoding(CharEncoding.UTF_8);
        response.addHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode("艺术生信息.xlsx", "UTF-8"));
        try {
            workbook.write(response.getOutputStream());
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private List<ArtCrowd> getDataList() {
        List<ArtCrowd> dataList = new ArrayList<>(1000);
        for (int i = 0; i < 1000; i++) {
            ArtCrowd artCrowd = new ArtCrowd();
            if (i % 2 == 0) {
                artCrowd.setName("Tom");
                artCrowd.setAge(19);
                artCrowd.setGender("Man");
                artCrowd.setPaintingLevel("一级证书");
                artCrowd.setDance(false);
                artCrowd.setAssessmentTime(LocalDateTime.now());
                artCrowd.setHobby("摔跤");
            } else {
                artCrowd.setName("Marry");
                artCrowd.setAge(18);
                artCrowd.setGender("Woman");
                artCrowd.setPaintingLevel("一级证书");
                artCrowd.setDance(true);
                artCrowd.setAssessmentTime(LocalDateTime.now());
                artCrowd.setHobby("钓鱼");
            }
            dataList.add(artCrowd);
        }
        return dataList;
    }
}
