package cn.iflytek.bid2tender;

import com.spire.doc.*;
import com.spire.doc.collections.ColumnCollection;
import com.spire.doc.collections.DocumentObjectCollection;
import com.spire.doc.collections.SectionCollection;
import com.spire.doc.documents.Paragraph;
import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Test {

    public static XWPFDocument readHandle(String filePath) {
        XWPFDocument resultDoc = null;
        try (FileInputStream fileInputStream = new FileInputStream(filePath);
             XWPFDocument doc = new XWPFDocument(fileInputStream);) {
            resultDoc = doc;


//            XWPFNum num = doc.getNumbering().getNum(new BigInteger("1"));
//            System.out.println(num.getCTNum());

            List<IBodyElement> bodyElements = doc.getBodyElements();
//            IBodyElement bodyElement = bodyElements.get(0);
//            System.out.println(bodyElement.getBody());
            Pattern pat = Pattern.compile("^目\\s{0,}录$");
            int index = 0;
            System.out.println(bodyElements.size());
            for (int i = 0; i < bodyElements.size(); i++) {
                IBodyElement ele = bodyElements.get(i);
                if (ele.getElementType() == BodyElementType.PARAGRAPH) {
                    index ++;
                }
//                System.out.println(ele.getBody().getBodyElements().size());
//                System.out.println(ele.getElementType());
//                XWPFParagraph para = (XWPFParagraph) ele;
//                System.out.println(para.getText());
//                Matcher matcher = pat.matcher(para.getText());
//                if (matcher.matches()) {
//                    index = i;
//                    break;
//                }
//                XWPFParagraph paragraph = resultDoc.createParagraph();
//                paragraph.setStyle(para.getStyleID());
            }
            System.out.println(index);
//            System.out.println("enc");
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
//        new XWPFDocument()

        return resultDoc;

    }

    public static void documentParse() {

    }

    public static void main(String[] args) throws IOException {
        String inPath = "D:/temp/提取/中国电信安徽公司2018年VoLTE网络端到端信令分析系统建设项目比选文件发售版/中国电信安徽公司2018年VoLTE网络端到端信令分析系统建设项目比选文件.docx";
        String outPath = "D:/temp/提取/中国电信安徽公司2018年VoLTE网络端到端信令分析系统建设项目比选文件发售版/test.docx";

        XWPFDocument doc = readHandle(inPath);


//        FileOutputStream out = new FileOutputStream(outPath);
//        doc.write(out);
//        out.close();

//        XWPFDocument doc = new XWPFDocument();
//        XWPFTable table = doc.createTable(4, 4);
//
//
//        doc.write(out);
//        out.close();

    }
}
