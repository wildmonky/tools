package org.example;

import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFPrintSetup;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jodconverter.core.document.DefaultDocumentFormatRegistry;
import org.jodconverter.core.office.OfficeException;
import org.jodconverter.core.office.OfficeUtils;
import org.jodconverter.local.LocalConverter;
import org.jodconverter.local.filter.PagesSelectorFilter;
import org.jodconverter.local.office.LocalOfficeManager;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPageSetup;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STOrientation;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args )
    {
        System.out.println( "Hello World!" );
    }

    public static void excelToImage() {

        File inputFile = new File("1.xlsx");
        File outputFile = new File("2.xlsx");
        File pdfFile = new File("3.pdf");
        File imageDir = new File("/imgDir");


        try (
                XSSFWorkbook workbook = new XSSFWorkbook(inputFile)
        ){
            XSSFSheet sheet = workbook.getSheetAt(0);

            // 设置打印区域
            workbook.setPrintArea(workbook.getSheetIndex(sheet), "$A$!:$B$10");
            workbook.setPrintArea(workbook.getSheetIndex(sheet), 0, 9, 0, 1);

            XSSFPrintSetup printSetup = sheet.getPrintSetup();
            printSetup.setTopMargin(0.05);
//            printSetup.setRightMargin(0.05);
//            printSetup.setBottomMargin(0.05);
            printSetup.setLeftMargin(0.05);

            CTWorksheet ctWorksheet = sheet.getCTWorksheet();
            CTPageSetup pageSetup = ctWorksheet.getPageSetup();
            // 设置纸张方向
            pageSetup.setOrientation(STOrientation.Enum.forString("portrait"));
            pageSetup.setPaperWidth("262mm");
            pageSetup.setPaperHeight("120mm");

            workbook.write(Files.newOutputStream(outputFile.toPath()));
        } catch (IOException | InvalidFormatException e) {
            throw new RuntimeException(e);
        }

        // Create an office manager using the default configuration.
        // The default port is 2002. Note that when an office manager
        // is installed, it will be the one used by default when
        // a converter is created.
        LocalOfficeManager officeManager = LocalOfficeManager.builder()
                .officeHome("D:\\Other_Tools\\LibreOffice\\24.8.2.1").install().build();
        try {
            // Start an office process and connect to the started instance (on port 2002).
            officeManager.start();

            PagesSelectorFilter filter = new PagesSelectorFilter(1);

            LocalConverter.builder()
                    .officeManager(officeManager)
//                    .loadDocumentMode(LoadDocumentMode.LOCAL)
                    .filterChain(filter)
                    .build()
                    .convert(outputFile)
                    .as(DefaultDocumentFormatRegistry.XLSX)
                    .to(pdfFile)
                    .as(DefaultDocumentFormatRegistry.PDF)
                    .execute();
        } catch (OfficeException e) {
            throw new RuntimeException(e);
        } finally {
            // Stop the office process
            OfficeUtils.stopQuietly(officeManager);
        }

        try (
                PDDocument document = Loader.loadPDF(outputFile)
        ){
            PDFRenderer pdfRenderer = new PDFRenderer(document);

            for (int pageIndex = 0; pageIndex < document.getNumberOfPages(); pageIndex++) {
                PDPage page = document.getPage(pageIndex);
                // 获取尺寸，单位点，1点 = 1/72英寸，可打印指定页
                float width = page.getMediaBox().getWidth();
                float height = page.getMediaBox().getHeight();
//                DPI（每英寸点数）是影响图像清晰度的关键参数。以下是一些常见的 DPI 设置及其效果：
//                  1、72 DPI：这是屏幕分辨率的标准，适合在网页上查看，但细节可能不够清晰。
//                  2、150 DPI：适合一般用途，比如在打印中保持相对清晰，适合阅读文档，但仍然不够细致。
//                  3、300 DPI：通常用于高质量打印，细节非常清晰，适合需要高精度的场合，比如打印和出版。
//                  4、600 DPI：适用于需要极高细节的情况，通常用于专业印刷，但文件大小会明显增加。
                BufferedImage bim = pdfRenderer.renderImageWithDPI(pageIndex, 300); // 设置DPI
                ImageIO.write(bim, "PNG", new File(imageDir + "page_" + (pageIndex + 1) + ".png"));
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }
}
