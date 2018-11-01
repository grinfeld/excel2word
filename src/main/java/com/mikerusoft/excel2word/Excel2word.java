package com.mikerusoft.excel2word;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.util.List;
import java.util.Map;
import java.util.UUID;

@SpringBootApplication
public class Excel2word implements CommandLineRunner {

    public static void main(String[] args){
        SpringApplication.run(Excel2word.class, args);
    }

    @Autowired DataReader dr;
    @Autowired WordProcessor<List<List<String>>> wr;
    @Autowired DataOutputer dt;

    @Override
    public void run(String... strings) throws Exception {
        List<Map<String, String>> data = dr.readData();
        List<List<String>> res = wr.processData(data);
        dt.output(res.get(0), UUID.randomUUID().toString() + ".docx");
    }
}
