package cn.iflytek.bid2tender;

import lombok.Data;

import java.util.List;

@Data
public class TenderFileInfo {
    private String fileName;
    private List<String> titles;
}
