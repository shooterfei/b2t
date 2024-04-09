package cn.iflytek.bid2tender;

import cn.iflytek.bid2tender.utils.ToolUtil;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.lang.annotation.ElementType;
import java.util.ArrayList;
import java.util.List;

public class PoiReadTest {
    public static void main(String[] args) throws IOException {

        File file = new File(Constant.baseFilePath);
        if (file.exists()) {
            if (file.isDirectory()) {
                File[] files = file.listFiles();
                for (File f : files) {
                    if (f.getName().contains("技术规范书") && f.getName().charAt(0) != '~' && f.isFile()) {
                        FileInputStream inputStream = new FileInputStream(Constant.baseFilePath + f.getName());
                        XWPFDocument xwpfDocument = new XWPFDocument(inputStream);
                        List<String[]> strings = techTitleExtract(xwpfDocument);
                        for (String[] strs : strings) {
                            for (String str : strs) {
                                System.out.print(str + "\t");
                            }
                            System.out.println();
                        }
                    }
                }
            }
        }

    }

    public static List<String[]> techTitleExtract(XWPFDocument document) {
        System.out.println(document.getSettings());
        List<IBodyElement> bodyElements = document.getBodyElements();
        XWPFStyles styles = document.getStyles();
        int level1 = 0, level2 = 0, count = 0;
        List<String[]> result = new ArrayList<>();

        for (int i = 0; i < bodyElements.size(); i++) {
            IBodyElement bodyElement = bodyElements.get(i);
            if (bodyElement.getElementType() == BodyElementType.PARAGRAPH) {
                XWPFParagraph paragraph = (XWPFParagraph) bodyElement;
                if (paragraph.getStyle() != null) {
                    XWPFStyle style = styles.getStyle(paragraph.getStyle());
//                    System.out.println(style.getName());
                    if ("heading 1".equals(style.getName())){
                        level2 = 0;
                        level1++;
                        count++;
                        String[] strings = new String[6];

                        strings[0] = String.valueOf(count);
                        strings[1] = ToolUtil.convertNumber(level1);
                        strings[2] = paragraph.getText();
                        strings[3] = "";
                        strings[4] = "满足、无偏离";
                        strings[5] = "";
                        result.add(strings);

                    }
                    if( "heading 2".equals(style.getName()))  {
                        level2++;
                        count++;

                        String[] strings = new String[6];
                        strings[0] = String.valueOf(count);
                        strings[1] = level1 + "." + level2;
                        strings[2] = paragraph.getText();
                        strings[3] = "";
                        strings[4] = "满足、无偏离";
                        strings[5] = "";
                        result.add(strings);
                    }
                }
            }
        }
        return result;
    }
}
