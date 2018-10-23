package com.mikerusoft.excel2word;

import java.util.List;
import java.util.Map;

public interface WordProcessor<Output> {
    Output processData(List<Map<String, String>> data);
}
