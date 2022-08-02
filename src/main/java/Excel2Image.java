import gui.ava.html.Html2Image;
import gui.ava.html.renderer.ImageRenderer;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.converter.ExcelToHtmlConverter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.w3c.dom.Document;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;

public class Excel2Image {
    public static void saveImage() {
        try {
            File file = new File("src/main/resources/报价模板.xls");
            String fileName = file.getName();
            FileInputStream inputStream = new FileInputStream(file);
            ByteArrayOutputStream outStream = new ByteArrayOutputStream();
            Document htmlDocument;
            Workbook workbook = null;
            HSSFWorkbook hSSFWorkbook = null;
            try {
                ExcelToHtmlConverter converter = new ExcelToHtmlConverter(DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());
                converter.setOutputColumnHeaders(false);
                converter.setOutputRowNumbers(false);
                //转化为 Workbook 对象
                workbook = WorkbookFactory.create(inputStream);
                for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                    String sheetName = StringUtils.leftPad(" ", i + 1, " ");
                    workbook.setSheetName(i, sheetName);
                }
                //判断Workbook 是XLSX，2003后版本的
                if (workbook instanceof XSSFWorkbook) {
                    XSSFWorkbook s = (XSSFWorkbook) workbook;
                    hSSFWorkbook = new HSSFWorkbook();
                    Xssf2Hssf xlsx2xls = new Xssf2Hssf();
                    xlsx2xls.transformXSSF(s, hSSFWorkbook);
                } else if (workbook instanceof HSSFWorkbook) {
                    //判断Workbook 是XLS
                    hSSFWorkbook = (HSSFWorkbook) workbook;
                } else {
                    throw new Exception("unknow class of workBook :" + workbook.getClass().getTypeName());
                }
                //转换html
                converter.processWorkbook(hSSFWorkbook);
                htmlDocument = converter.getDocument();
                DOMSource domSource = new DOMSource(htmlDocument);
                StreamResult streamResult = new StreamResult(outStream);
                TransformerFactory tfFactory = TransformerFactory.newInstance();
                Transformer tf = tfFactory.newTransformer();
                tf.setOutputProperty(OutputKeys.INDENT, "yes");
                tf.setOutputProperty(OutputKeys.METHOD, "html");
                tf.transform(domSource, streamResult);
                Html2Image image = Html2Image.fromHtml(outStream.toString());
                ImageRenderer imageRenderer = image.getImageRenderer();
                imageRenderer.saveImage(new File("src/main/resources/" + fileName + ".png"));
            } catch (Exception e3) {
            } finally {
                outStream.close();
                if (workbook != null) {
                    workbook.close();
                }
                if (hSSFWorkbook != null) {
                    hSSFWorkbook.close();
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {
        saveImage();
    }

}
