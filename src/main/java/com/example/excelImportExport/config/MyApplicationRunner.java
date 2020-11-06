package com.example.excelImportExport.config;


import lombok.extern.slf4j.Slf4j;
import org.springframework.boot.ApplicationArguments;
import org.springframework.boot.ApplicationRunner;
import org.springframework.core.annotation.Order;
import org.springframework.core.env.Environment;
import org.springframework.stereotype.Component;

import javax.annotation.Resource;
import java.net.InetAddress;

@Slf4j
@Component
@Order(value = 1)
public class MyApplicationRunner implements ApplicationRunner {

    @Resource
    private Environment environment;

    /**
     * 工程启动结束后，立即打开swagger页面
     *
     * @param args -
     */
    @Override
    public void run(ApplicationArguments args) {

        try {
            String url = "http://" + InetAddress.getLocalHost().getHostAddress() + ":" + environment.getProperty("server.port") + "/swagger-ui.html";
            log.info("访问链接：" + url);

            Runtime.getRuntime().exec("cmd /c start " + url);
        } catch (Exception e) {
            e.printStackTrace();
        }

    }
}
