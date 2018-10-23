package com.mikerusoft.excel2word;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.util.List;
import java.util.Map;

@SpringBootApplication
public class Excel2word implements CommandLineRunner {

    public static void main(String[] args){
        SpringApplication.run(Excel2word.class, args);
    }

    @Autowired DataReader dr;

    @Override
    public void run(String... strings) throws Exception {
        List<Map<String, String>> data = dr.readData();
        WordProcessor<String> wr = new ReplaceWordProcessor();
        String res = wr.processData(data);
        DataOutputer dt = new FileDataOutputer();
        dt.output();
    }
}
