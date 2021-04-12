package com.xzy1024.demo.utils;

import com.aspose.cells.Workbook;
import com.aspose.slides.Presentation;
import com.aspose.slides.SaveFormat;
import com.aspose.words.License;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Image;
import com.itextpdf.text.pdf.PdfWriter;

import javax.servlet.http.HttpServletRequest;
import java.io.*;
import java.util.ArrayList;

/**
 * 文档类型转换为PDF预览
 *
 * @author Mr.Wang
 * @date 2021/4/2 14:02
 */
public class PDFUtil {

    //word转换xml
    private static String wordLicense = "<License><Data><Products><Product>Aspose.Total for Java</Product><Product>Aspose.Words for Java</Product></Products><EditionType>Enterprise</EditionType><SubscriptionExpiry>20991231</SubscriptionExpiry><LicenseExpiry>20991231</LicenseExpiry><SerialNumber>8bfe198c-7f0c-4ef8-8ff0-acc3237bf0d7</SerialNumber></Data><Signature>sNLLKGMUdF0r8O1kKilWAGdgfs2BvJb/2Xp8p5iuDVfZXmhppo+d0Ran1P9TKdjV4ABwAgKXxJ3jcQTqE/2IRfqwnPf8itN8aFZlV3TJPYeD3yWE7IT55Gz6EijUpC7aKeoohTb4w2fpox58wWoF3SNp6sK6jDfiAUGEHYJ9pjU=</Signature></License>";


    //判断文档类型
    public static boolean convert2PDF(String intPath, String outPath) {
        String suffix = getFileSufix(intPath);
        File file = new File(intPath);
        if (!file.exists()) {
            return false;
        }
        if (suffix.equals("pdf")) {
            return false;
        }
        if (suffix.equals("doc") || suffix.equals("docx")) {
            return word2PDF(intPath, outPath);
        } else if (suffix.equals("ppt") || suffix.equals("pptx")) {
            return ppt2PDF(intPath, outPath);
        } else if (suffix.equals("xls") || suffix.equals("xlsx")) {
            return excel2PDF(intPath, outPath);
        } else {
            return false;
        }
    }

    public static String getFileSufix(String fileName) {
        int splitIndex = fileName.lastIndexOf(".");
        return fileName.substring(splitIndex + 1);
    }

    // word转换为pdf
    public static boolean word2PDF(String inputFile, String outPath) {
        try {
//            InputStream stream = PDFUtil.class.getClassLoader()
//                    .getResourceAsStream("license.xml");
            ByteArrayInputStream stream = new ByteArrayInputStream(wordLicense.getBytes());
            License aposeLic = new License();
            aposeLic.setLicense(stream);
            //doc
            com.aspose.words.Document document = new com.aspose.words.Document(inputFile);
            document.save(outPath, com.aspose.words.SaveFormat.PDF);
            return true;
        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }
    }

    // excel转换为pdf
    public static boolean excel2PDF(String inputFile, String outPath) {
        try {

//            InputStream stream = PDFUtil.class.getClassLoader()
//                    .getResourceAsStream("license.xml");
            ByteArrayInputStream stream = new ByteArrayInputStream(wordLicense.getBytes());
            com.aspose.cells.License aposeLic = new com.aspose.cells.License();
            aposeLic.setLicense(stream);
            File excelFile = new File(outPath);
            Workbook workBook = new Workbook(inputFile);
            FileOutputStream fileOs = new FileOutputStream(excelFile);
            workBook.save(fileOs, com.aspose.cells.SaveFormat.PDF);
            fileOs.close();
            return true;
//            Evaluation Only. Created with Aspose.Cells for Java. Copyright 2003 - 2021 Aspose Pty Ltd.
        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }
    }

    // ppt转换为pdf
    public static boolean ppt2PDF(String inputFile, String outPath) {
        try {
//            InputStream stream = PDFUtil.class.getClassLoader()
//                    .getResourceAsStream("license.xml");
            ByteArrayInputStream stream = new ByteArrayInputStream(wordLicense.getBytes());
            com.aspose.slides.License aposeLic = new com.aspose.slides.License();
            aposeLic.setLicense(stream);
            File file = new File(outPath);// 输出pdf路径
            Presentation pres = new Presentation(inputFile);//输入ppt路径
            FileOutputStream fileOS = new FileOutputStream(file);
            //IFontsManager fontsManager = pres.getFontsManager();
            pres.save(fileOS, SaveFormat.Pdf);
            fileOS.close();
            return true;
        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }
    }


    public static boolean imgOfPdf(String inputFilePath, String outFilePath) {
        boolean result = false;
        String pdfUrl = "";
        String fileUrl = "";
        try {
            ArrayList<String> imageUrllist = new ArrayList<String>(); // 图片list集合
            // 添加图片文件路径
            imageUrllist.add(inputFilePath);
            result = true;
            if (result == true) {
                File file = PDFUtil.Pdf(imageUrllist, outFilePath);// 生成pdf
                file.createNewFile();
            }
            return true;
        } catch (IOException e) {
            e.printStackTrace();
            return false;
        }

    }

    public static File Pdf(ArrayList<String> imageUrllist,
                           String mOutputPdfFileName) {
        //Document doc = new Document(PageSize.A4, 20, 20, 20, 20); // new一个pdf文档
        com.itextpdf.text.Document doc = new com.itextpdf.text.Document();
        try {

            PdfWriter
                    .getInstance(doc, new FileOutputStream(mOutputPdfFileName)); // pdf写入
            doc.open();// 打开文档
            for (int i = 0; i < imageUrllist.size(); i++) { // 循环图片List，将图片加入到pdf中
                doc.newPage(); // 在pdf创建一页
                Image png1 = Image.getInstance(imageUrllist.get(i)); // 通过文件路径获取image
                float heigth = png1.getHeight();
                float width = png1.getWidth();
                int percent = getPercent2(heigth, width);
                png1.setAlignment(Image.MIDDLE);
                png1.scalePercent(percent + 3);// 表示是原来图像的比例;
                doc.add(png1);
            }
            doc.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (DocumentException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        File mOutputPdfFile = new File(mOutputPdfFileName); // 输出流
        if (!mOutputPdfFile.exists()) {
            mOutputPdfFile.deleteOnExit();
            return null;
        }
        return mOutputPdfFile; // 反回文件输出流
    }

    public static int getPercent2(float h, float w) {
        int p = 0;
        float p2 = 0.0f;
        p2 = 530 / w * 100;
        p = Math.round(p2);
        return p;
    }


}
