package com.ncsi;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.autoconfigure.r2dbc.R2dbcAutoConfiguration;
import org.springframework.boot.autoconfigure.data.r2dbc.R2dbcDataAutoConfiguration;
import org.springframework.boot.autoconfigure.data.r2dbc.R2dbcRepositoriesAutoConfiguration;
import org.springframework.context.annotation.ComponentScan;
import org.springframework.boot.autoconfigure.domain.EntityScan;
import org.springframework.data.jpa.repository.config.EnableJpaRepositories;

@SpringBootApplication(
    exclude = {
        R2dbcAutoConfiguration.class,
        R2dbcDataAutoConfiguration.class,
        R2dbcRepositoriesAutoConfiguration.class
    },
    scanBasePackages = {"com.ncsi"}
)
@EntityScan("com.ncsi")
@EnableJpaRepositories("com.ncsi")
@ComponentScan(basePackages = "com.ncsi")
public class NcsiApplication {
    public static void main(String[] args) {
        System.setProperty("spring.autoconfigure.exclude", 
            "org.springframework.boot.autoconfigure.r2dbc.R2dbcAutoConfiguration," +
            "org.springframework.boot.autoconfigure.data.r2dbc.R2dbcDataAutoConfiguration," +
            "org.springframework.boot.autoconfigure.data.r2dbc.R2dbcRepositoriesAutoConfiguration");
        SpringApplication.run(NcsiApplication.class, args);
    }
} 