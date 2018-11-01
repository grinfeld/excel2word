package com.mikerusoft.excel2word.props.word;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.context.annotation.Configuration;

@Configuration
@ConfigurationProperties("word")
@Data
@NoArgsConstructor
@AllArgsConstructor
public class WordProperties {
    private String prefix;
    private String[] templates;
}
