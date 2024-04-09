package cn.iflytek.bid2tender;

import cn.iflytek.bid2tender.utils.ToolUtil;
import com.spire.doc.*;
import com.spire.doc.collections.*;
import com.spire.doc.documents.*;
import com.spire.doc.interfaces.IStyle;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.tools.ant.types.resources.Difference;

import javax.management.remote.rmi._RMIConnection_Stub;
import javax.security.auth.Subject;
import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class SpireTest {

    public static List<DocumentObject> getSubjects(Document doc) {
        ArrayList<DocumentObject> documentObjects = new ArrayList<>();
        DocumentObjectCollection c1 = doc.getChildObjects();
        for (int i = 0; i < c1.getCount(); i++) {
            DocumentObject d1 = c1.get(i);
            DocumentObjectCollection c2 = d1.getChildObjects();
            for (int j = 0; j < c2.getCount(); j++) {
                DocumentObject d2 = c2.get(j);
//                if (d2.getDocumentObjectType() == DocumentObjectType.Paragraph || d2.getDocumentObjectType() == DocumentObjectType.Table) {
//                    documentObjects.add(d2);
//                }
                DocumentObjectCollection c3 = d2.getChildObjects();
                for (int k = 0; k < c3.getCount(); k++) {
                    DocumentObject d3 = c3.get(k);
                    if (d3.getDocumentObjectType() == DocumentObjectType.Paragraph || d3.getDocumentObjectType() == DocumentObjectType.Table) {
                        documentObjects.add(d3);
                    }
                }
            }
        }
        return documentObjects;
    }

    public static List<TenderFileInfo> titleExtract(Document doc) {
        // todo 前附表中提取标题信息
        Pattern startPattern = Pattern.compile("^参选人须知前附表$");
        Pattern electionCompositionPattern = Pattern.compile("^参选文件组成$");


        // 文件名抽取正则
        Pattern fielNameExtractPattern = Pattern.compile("^\\d.\\d(.*参选文件).*$");
        // 标题提取正则
        Pattern titleExtractPattern = Pattern.compile("^(.*?)([（；;]).*$");
        List<DocumentObject> subjects = getSubjects(doc);
        boolean flag = false;
        ArrayList<TenderFileInfo> tenderFileInfos = new ArrayList<>();
        for (int i = 0; i < subjects.size(); i++) {
            DocumentObject subject = subjects.get(i);
            if (subject.getDocumentObjectType() == DocumentObjectType.Paragraph) {
                Paragraph para = (Paragraph) subject;
                String text = para.getText();
                Matcher matcher = startPattern.matcher(text);
                if (matcher.matches()) {
                    flag = true;
                }
            } else if (subject.getDocumentObjectType() == DocumentObjectType.Table) {
                if (flag) {
                    Table tab = (Table) subject;
                    RowCollection rows = tab.getRows();
                    TableCell cell = null;
                    rows:
                    for (int j = 0; j < rows.getCount(); j++) {
                        TableRow row = rows.get(j);
                        CellCollection cells = row.getCells();

                        int tIndex = 0;
                        if (cells.getCount() == 3) {
                            tIndex = 1;
                        }
                        cell = cells.get(tIndex);
                        DocumentObjectCollection contents = cell.getChildObjects();
                        for (int k = 0; k < contents.getCount(); k++) {
                            DocumentObject dot = contents.get(k);
                            if (dot.getDocumentObjectType() == DocumentObjectType.Paragraph) {
                                Paragraph para = (Paragraph) dot;
                                String text = para.getText();
                                Matcher matcher = electionCompositionPattern.matcher(text);
                                if (matcher.matches()) {
                                    cell = cells.get(tIndex + 1);
                                    break rows;
                                }
                            }
                        }
                    }

                    DocumentObjectCollection contents = cell.getChildObjects();
                    boolean fileFlag = false;
                    for (int k = 0; k < contents.getCount(); k++) {
                        DocumentObject dot = contents.get(k);
                        if (dot.getDocumentObjectType() == DocumentObjectType.Paragraph) {
                            Paragraph para = (Paragraph) dot;
                            String text = para.getText();
                            if (para.getListFormat().getListType() == ListType.No_List) {
                                Matcher matcher = fielNameExtractPattern.matcher(text);
                                if (matcher.matches()) {
                                    fileFlag = true;
                                    TenderFileInfo tenderFileInfo = new TenderFileInfo();
                                    if ("电子参选文件".equals(matcher.group(1))) {
                                        fileFlag = false;
                                    } else {
                                        tenderFileInfo.setFileName(matcher.group(1));
                                        tenderFileInfos.add(tenderFileInfo);
                                        ArrayList<String> titles = new ArrayList<>();
                                        tenderFileInfo.setTitles(titles);
                                    }
                                }
//                                System.out.println(text);
                            } else if (para.getListFormat().getListType() == ListType.Numbered) {
                                if (fileFlag) {
                                    int size = tenderFileInfos.size();
                                    Matcher matcher = titleExtractPattern.matcher(text);
                                    if (matcher.matches()) {
                                        tenderFileInfos.get(size - 1).getTitles().add(matcher.group(1));
                                    }
                                }
                            }
//                            System.out.println(para.getListFormat().getListType());
////                            System.out.println(para);
//                            System.out.println("11111111111111111111111");
                        }
                    }


                    flag = false;
                }
            }
        }
        return tenderFileInfos;
    }

    public static List<DocumentObject> tableInfoExtract(Document doc) {
        List<DocumentObject> subjects = getSubjects(doc);
        Pattern pattern = Pattern.compile("^\\s{0,}参选人须知前附表\\s{0,}$");
        Pattern titlePattern = Pattern.compile("^(Heading|Title).*$");
        boolean flag = false;
        int[] tableRangeIndex = new int[2];
        for (int i = 0; i < subjects.size(); i++) {
            DocumentObject subject = subjects.get(i);
            if (subject.getDocumentObjectType() == DocumentObjectType.Paragraph) {
                Paragraph para = (Paragraph) subject;
                String text = para.getText();
                Matcher matcher = pattern.matcher(text);
                if (flag) {
                    String styleName = para.getStyle().getName();
                    if ("评选索引表".equals(text)) {
                        tableRangeIndex[0] = i;
//                            System.out.println(styleName + " : " + text);
                    }
                    Matcher titleMatcher = titlePattern.matcher(styleName);
                    if (titleMatcher.matches()) {
                        if (text.contains("专用印章授权函")) {
                            tableRangeIndex[1] = i - 1;
                            break;
                        }
                    }
                }
                if (matcher.matches()) {
                    flag = true;
                }
            }
        }
        return subjects.subList(tableRangeIndex[0], tableRangeIndex[1]);
    }

    public static void addDocBodyList(Document doc, List<DocumentObject> bodys) {
        DocumentObjectCollection tdc = doc.getSections().getLastItem().getFirstChild().getChildObjects();
        for (DocumentObject documentObject : bodys) {
            tdc.add(documentObject.deepClone());
        }
    }

    public static void addDocBody(Document doc, DocumentObject body) {
        DocumentObjectCollection tdc = doc.getSections().getLastItem().getFirstChild().getChildObjects();
        tdc.add(body.deepClone());
    }


    public static void addTitle(Document doc, String titleName, String level) {
        Paragraph paragraph = new Paragraph(doc);
        paragraph.setText(titleName);
        paragraph.applyStyle(level);
        Section sec = (Section) doc.getSections().getLastItem();
        sec.getBody().getChildObjects().add(paragraph);
    }

    /**
     * 添加标题样式
     *
     * @param doc
     */
    public static void addStyles(Document doc) {
        ParagraphStyle heading1 = new ParagraphStyle(doc);
        heading1.setName("Heading 1");
        heading1.getCharacterFormat().setFontNameBidi("Times New Roman");
        heading1.getCharacterFormat().setFontNameFarEast("宋体");
        heading1.getCharacterFormat().setFontNameNonFarEast("Times New Roman");
        heading1.getCharacterFormat().setFontNameAscii("Times New Roman");
        heading1.getCharacterFormat().setFontSize(22.0f);
        heading1.getCharacterFormat().setBold(true);
        heading1.getCharacterFormat().setBoldBidi(true);
        heading1.getCharacterFormat().setFontSizeBidi(22.0f);

        heading1.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Justify);
        heading1.getParagraphFormat().setKeepLines(true);
        heading1.getParagraphFormat().setKeepFollow(true);
        heading1.getParagraphFormat().setBeforeSpacing(17.0f);
        heading1.getParagraphFormat().setAfterSpacing(16.5f);
        heading1.getParagraphFormat().isWidowControl(false);
        heading1.getParagraphFormat().setLineSpacing(28.9f);
        heading1.getParagraphFormat().setLineSpacingRule(LineSpacingRule.Multiple);

        heading1.applyBaseStyle("Normal");

        ParagraphStyle heading2 = new ParagraphStyle(doc);
        heading2.setName("Heading 2");
        heading2.getCharacterFormat().setFontNameBidi("Times New Roman");
        heading2.getCharacterFormat().setFontNameFarEast("宋体");
        heading2.getCharacterFormat().setFontNameNonFarEast("Calibri Light");
        heading2.getCharacterFormat().setFontNameAscii("Calibri Light");
        heading2.getCharacterFormat().setFontSize(16.0f);
        heading2.getCharacterFormat().setBold(true);
        heading2.getCharacterFormat().setBoldBidi(true);
        heading2.getCharacterFormat().setFontSizeBidi(16.0f);

        heading2.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Justify);
        heading2.getParagraphFormat().setKeepLines(true);
        heading2.getParagraphFormat().setKeepFollow(true);
        heading2.getParagraphFormat().setBeforeSpacing(13.0f);
        heading2.getParagraphFormat().setAfterSpacing(13.0f);
        heading2.getParagraphFormat().isWidowControl(false);
        heading2.getParagraphFormat().setLineSpacing(20.8f);
        heading2.getParagraphFormat().setLineSpacingRule(LineSpacingRule.Multiple);

        heading2.applyBaseStyle("Normal");

        doc.getStyles().add(heading1);
        doc.getStyles().add(heading2);
    }


    /**
     * 添加换页
     *
     * @param doc
     */
    public static void addPageBreak(Document doc) {
        Paragraph paragraph = new Paragraph(doc);
        paragraph.appendBreak(BreakType.Page_Break);
        Section sec = (Section) doc.getSections().getLastItem();
        sec.getBody().getChildObjects().add(paragraph);
    }

    /**
     * 提取相应title中的具体信息
     *
     * @param doc
     * @param title
     * @return
     */
    private static List<DocumentObject> extractTitleInfo(Document doc, String title) {
        List<DocumentObject> subjects = getSubjects(doc);
        Pattern comparePattern = Pattern.compile("^(\\d+\\.){0,}(.*)$");
        // 资质 | 资格 两字可替换使用, 特殊处理文本
        boolean aptitudeFlag = false;
        if (title.contains("资质") || title.contains("资格")) {
            aptitudeFlag = true;
        }
        int index = -1;
        boolean eFlag = false;
        for (int i = 0; i < subjects.size(); i++) {
            DocumentObject subject = subjects.get(i);
            if (subject.getDocumentObjectType() == DocumentObjectType.Paragraph) {
                Paragraph paragraph = (Paragraph) subject;
                String text = paragraph.getText();
                if (title.equals(text)) {
                    eFlag = true;
                    index = i;
                } else {
                    Matcher matcher = comparePattern.matcher(text);
                    if (matcher.matches()) {
                        if (aptitudeFlag) {
                            String textRepl = matcher.group(2).replace("资格", "资质");
                            String titleRepl = title.replace("资格", "资质");
                            if (textRepl.equals(titleRepl)) {
                                eFlag = false;
                                index = i;
                            }
                        } else {
                            if (matcher.group(2).equals(title)) {
                                eFlag = false;
                                index = i;
                            }
                        }
                    }
                }
            }
        }
        if (index > 0) {
            System.out.println(index + " : " + title + " : " + ((Paragraph) subjects.get(index)).getText());
            ArrayList<DocumentObject> documentObjects = new ArrayList<>();
            if (eFlag) documentObjects.add(subjects.get(index));
            for (int i = index + 1; i < subjects.size(); i++) {
                DocumentObject subject = subjects.get(i);
                if (subject.getDocumentObjectType() == DocumentObjectType.Paragraph) {
                    Paragraph paragraph = (Paragraph) subject;
                    if (!paragraph.getStyle().getName().startsWith("Normal")) {
                        break;
                    }
                }
                documentObjects.add(subject);
            }
            return documentObjects;
        }

        return null;
    }


    public static void fileGenerate() {
        Document document = new Document(Constant.bidFilePath);
        List<TenderFileInfo> tenderFileInfos = titleExtract(document);
        for (int i = 0; i < tenderFileInfos.size(); i++) {
            TenderFileInfo tenderFileInfo = tenderFileInfos.get(i);
            Document target = new Document();
            // 封面拷贝
            SectionCollection sections = document.getSections();

            // 写入封面
            Section sec = sections.get(0).deepClone();
            target.getSections().add(sec);
            // 换页
            addPageBreak(target);

            // 生成样式
            addStyles(target);

            // 生成标题相关
            for (int j = 0; j < tenderFileInfo.getTitles().size(); j++) {
                String title = tenderFileInfo.getTitles().get(j);
                List<DocumentObject> documentObjects = extractTitleInfo(document, title);
                title = ToolUtil.convertNumber(j + 1) + "、" + title;
                addTitle(target, title, "Heading 1");
                if (documentObjects != null) {
                    addDocBodyList(target, documentObjects);
                    addPageBreak(target);
                }
            }
            if (tenderFileInfo.getFileName().contains("技术")) {
                SectionCollection secs = target.getSections();
                boolean flag = false;
                Table t = null;

                ft:
                for (int j = 0; j < secs.getCount(); j++) {
                    Section s = secs.get(j);
                    DocumentObjectCollection co = s.getBody().getChildObjects();
                    for (int k = 0; k < co.getCount(); k++) {
                        DocumentObject dot = co.get(k);
                        if (flag) {
                            if (dot.getDocumentObjectType() == DocumentObjectType.Table) {
                                t = (Table) dot;
                                break ft;
                            }
                        }
                        if (dot.getDocumentObjectType() == DocumentObjectType.Paragraph) {
                            Paragraph p = (Paragraph) dot;
                            if (p.getText().contains("采购需求偏离表")) {
                                flag = true;
                            }
                        }
                    }
                }
                if (t != null) {
                    File file = new File(Constant.baseFilePath);
                    if (file.exists()) {
                        if (file.isDirectory()) {
                            File[] files = file.listFiles();
                            for (File f : files) {
                                if (f.getName().contains("技术规范书") && f.getName().charAt(0) != '~' && f.isFile()) {
                                    try (FileInputStream inputStream = new FileInputStream(Constant.baseFilePath + f.getName());) {
                                        XWPFDocument xwpfDocument = new XWPFDocument(inputStream);
                                        List<String[]> tsl = PoiReadTest.techTitleExtract(xwpfDocument);
                                        RowCollection rows = t.getRows();
                                        if (rows.getCount() > (tsl.size() + 1)) {
                                            int differenceValue = rows.getCount() - tsl.size();
                                            for (int j = 0; j < differenceValue; j++) {
                                                rows.removeAt(rows.getCount() - 1);
                                            }
                                        }
                                        if (rows.getCount() <= tsl.size()) {
                                            int differenceValue = tsl.size() - rows.getCount();
                                            for (int j = 0; j <= differenceValue; j++) {
                                                TableRow tableRow = rows.get(rows.getCount() - 1);
                                                rows.add(tableRow.deepClone());
                                            }
                                        }
                                        for (int j = 1; j < rows.getCount(); j++) {
                                            TableRow tableRow = rows.get(j);
                                            CellCollection cells = tableRow.getCells();
                                            for (int k = 0; k < tsl.get(j - 1).length; k++) {
                                                cells.get(k).getFirstParagraph().setText(tsl.get(j - 1)[k]);
                                            }
                                        }
                                    } catch (IOException e) {
                                        throw new RuntimeException(e);
                                    }
                                }
                            }
                        }
                    }
                }
            }


            target.saveToFile(Constant.baseFilePath + "target/" + tenderFileInfo.getFileName() + ".docx", FileFormat.Docx);

        }
        System.out.println(tenderFileInfos);
    }

    public static List<String> techTitleExtract(Document doc) {
/*
        SectionCollection sections = doc.getSections();
        int ct1 = sections.getCount();
        while (ct1 > 0) {
            ct1 = sections.getCount();
            System.out.println(ct1);
            for (int i = 0; i < ct1; i++) {
                Section section = sections.get(i);
                DocumentObjectCollection cos = section.getBody().getChildObjects();
                int ct2 = cos.getCount();
                if (ct2 == 0) {
                    sections.removeAt(i);
                    break;
                }
                for (int j = 0; j < ct2; j++) {
                    DocumentObject r = cos.get(0);
                    if (r.getDocumentObjectType() == DocumentObjectType.Paragraph) {
                        Paragraph paragraph = (Paragraph) r;
//                        System.out.println(paragraph.getText());
                    }
                    cos.removeAt(0);
                }
            }
        }
*/
/*

        int index = 0;
        Document document = null;
        String filePath = null;
        SectionCollection sections = doc.getSections();
//        process:
//        while (true) {
        document = new Document();
        filePath = Constant.baseFilePath + "process/" + index + ".docx";
        StyleCollection styles = doc.getStyles();
        for (int i = 0; i < styles.getCount(); i++) {
            document.getStyles().add(styles.get(i).deepClone());
        }

//            doc.getStyles().a
        int sCount = sections.getCount();
        for (int i = 0; i < sCount; i++) {
            Section section = sections.get(i);
            Section sec = doc.addSection();
            DocumentObjectCollection cos = section.getChildObjects();
            for (int j = 0; j < cos.getCount(); j++) {
                DocumentObject dot = cos.get(j);
                if (dot.getDocumentObjectType() == DocumentObjectType.Body) {
                    DocumentObjectCollection childObjects = section.getBody().getChildObjects();
                    int count = childObjects.getCount();
                    if (count == 0) {
//                    document = null;
//                    filePath = null;
//                    break process;
                    }
                    for (int k = 0; k < count; k++) {
//                    System.out.println(j);
                        DocumentObject documentObject = childObjects.get(0);
                        if (documentObject.getDocumentObjectType() == DocumentObjectType.Paragraph) {
                            Paragraph paragraph = (Paragraph) documentObject;
                            System.out.println(paragraph.getText());
                        }
                        sec.getBody().getChildObjects().add(documentObject.deepClone());
                        section.getBody().getChildObjects().removeAt(0);
                    }
//                for (int j = 0; j < count; j++) {
//                }
//            System.out.println(childObjects.getCount());

                    ParagraphCollection paragraphs = section.getBody().getParagraphs();
                    System.out.println(paragraphs.getCount());
                    for (int k = 0; k < paragraphs.getCount(); k++) {
                        Paragraph paragraph = paragraphs.get(k);
                        String styleName = paragraph.getStyle().getName();
                        String text = paragraph.getText();
//                if ("Heading 1".equals(styleName) || "Heading 2".equals(styleName)) {
//                    System.out.println(styleName + " : " + text);
//                }
//                if (text.contains("VOLTE 网络业务质量")) {
//                    System.out.println("-----------");
//                }
                    }
                } else {
                    sec.getChildObjects().add(dot.deepClone());
                }
            }


        }
        index++;
//        }
        if (document != null) {
            document.saveToFile(filePath, FileFormat.Docx);
        }
*/
        return null;
    }

    public static void main(String[] args) throws FileNotFoundException {
        fileGenerate();
/*
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
*/
    }

}
