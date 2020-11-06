package com.example.excelImportExport.config;

import org.springframework.beans.BeansException;
import org.springframework.context.ApplicationContext;
import org.springframework.context.ApplicationContextAware;
import org.springframework.context.annotation.Configuration;

/**
 * 获取上线文
 */
@Configuration
public class SpringContext implements ApplicationContextAware {

    private static ApplicationContext applicationContext = null;

    @Override
    public void setApplicationContext(ApplicationContext applicationContext) throws BeansException {
        if (SpringContext.applicationContext == null) {
            synchronized (this) {
                if (SpringContext.applicationContext == null) {
                    SpringContext.applicationContext = applicationContext;
                }
            }
        }
    }

    /**
     * 获取对象，如service
     *
     * @param clazz -
     * @param <T>   -
     */
    public static <T> T getBean(Class<T> clazz) {
        return applicationContext.getBean(clazz);
    }
}
