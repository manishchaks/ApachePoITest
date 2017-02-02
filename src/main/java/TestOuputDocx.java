import java.io.*;

import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.*;


public class TestOuputDocx {

    private static final String DOCX_FILE = "borderdemo.docx";

    public static void main(String[] args) throws IOException {
        //Blank Document
        XWPFDocument document = new XWPFDocument();
        // Create output stream for file
        FileOutputStream out = new FileOutputStream(new File(DOCX_FILE));
        // write some text with border
        writeTextWithBorder(document, out);

        // open the file again and extract the text
        extractText();
    }

    private static void extractText() throws IOException {
        /// Extract text
        XWPFDocument docx = new XWPFDocument(new FileInputStream(DOCX_FILE));
        //using XWPFWordExtractor Class
        XWPFWordExtractor we = new XWPFWordExtractor(docx);
        System.out.println(we.getText());
    }

    private static void writeTextWithBorder(XWPFDocument document, FileOutputStream out) throws IOException {
        //create paragraph
        XWPFParagraph paragraph = document.createParagraph();

        //Set bottom border to paragraph
        paragraph.setBorderBottom(Borders.BASIC_BLACK_DASHES);

        //Set left border to paragraph
        paragraph.setBorderLeft(Borders.BASIC_BLACK_DASHES);

        //Set right border to paragraph
        paragraph.setBorderRight(Borders.BASIC_BLACK_DASHES);

        //Set top border to paragraph
        paragraph.setBorderTop(Borders.BASIC_BLACK_DASHES);

        XWPFRun run = paragraph.createRun();
        run.setText("With over two decades of technology and business experience, Pramati founders have had the unique opportunity of being part of the Indian IT right from its inception.\n" +
                "\n" +
                "Pramati was founded in 1998 by Jay and Vijay Pullur in Hyderabad, India. Founded as a web technology company, Pramati cut its teeth in enterprise-class web infrastructure technology. The Pramati Application Server was its flagship product that hit the market along with offerings from BEA Systems, IBM and Oracle. Initially adopted by the rapidly growing financial services industry in India, by 2001 the Pramati app server was installed by each of India's top 10 banks.\n" +
                "\n" +
                "Building on this early success, Pramati started to invest in a number of critical enterprise technologies, going on to become one of India's leading software product companies. The company has carved out a rare and unique path in India's impressive technological landscape - as a technology and product innovator.\n" +
                "\n" +
                "Pramati invests in cutting-edge technologies and people to create independent companies. The company employs consulting and delivery teams across three global locations.");


        // Write a paragraph with formatting
        //create paragraph
        XWPFParagraph secondParagraph = document.createParagraph();

        //Set Bold an Italic
        XWPFRun paragraphOneRunOne = secondParagraph.createRun();
        paragraphOneRunOne.setBold(true);
        paragraphOneRunOne.setItalic(true);
        paragraphOneRunOne.setText("Font Style");
        paragraphOneRunOne.addBreak();

        //Set text Position
        XWPFRun paragraphOneRunTwo = secondParagraph.createRun();
        paragraphOneRunTwo.setText("Font Style two");
        paragraphOneRunTwo.setTextPosition(100);

        //Set Strike through and Font Size and Subscript
        XWPFRun paragraphOneRunThree = secondParagraph.createRun();
        paragraphOneRunThree.setStrike(true);
        paragraphOneRunThree.setFontSize(20);
        paragraphOneRunThree.setSubscript(VerticalAlign.SUBSCRIPT);
        paragraphOneRunThree.setText(" Different Font Styles");


        document.write(out);
        out.close();

        System.out.println("applyingborder.docx written successully");
    }
}