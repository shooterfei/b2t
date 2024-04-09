package cn.iflytek.bid2tender;

import com.aspose.words.*;

import java.io.File;



public class AsposeTest {


    public static  void techTitleExtract (Document doc) {
        NodeCollection childNodes = doc.getSections().get(1).getBody().getChildNodes(NodeType.PARAGRAPH, false);
        for (int i = 0; i < childNodes.getCount(); i++) {
            Node r = childNodes.get(i);
            if (r.getNodeType() == NodeType.PARAGRAPH) {
                Paragraph p = (Paragraph) r;
                System.out.println(p.getText());
            }
        }
    }

    public static void main(String[] args) throws Exception {

        File file = new File(Constant.baseFilePath);
        if (file.exists()) {
            if (file.isDirectory()) {
                File[] files = file.listFiles();
                for (File f : files) {
                    if (f.getName().contains("技术规范书") && f.getName().charAt(0) != '~' && f.isFile()) {
                        Document document = new Document(Constant.baseFilePath + f.getName());
                        techTitleExtract(document);
                    }
                }
            }
        }
    }
}
