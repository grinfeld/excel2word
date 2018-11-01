package com.mikerusoft.excel2word.props.excel;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.context.annotation.Configuration;

import java.util.Optional;

@Configuration
@ConfigurationProperties("excel")
@Data
@NoArgsConstructor
@AllArgsConstructor
public class ExcelProperties {

    private Sheet sheet;
    private File file;

    public Optional<Sheet> getSheetOptional() { return Optional.ofNullable(sheet); }
    public Optional<File> getFileOptional() { return Optional.ofNullable(file); }

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    public static class Sheet {
        private String name;
    }

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    public static class File {
        private String name;
    }
}
