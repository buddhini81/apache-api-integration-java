package com.ourcompany.app;

import com.ourcompany.app.data.QuoteItem;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Locale;

public class WordQuotation {
    public void generateQuotation(List<QuoteItem> quoteData) {
        XWPFDocument document = new XWPFDocument();
        try(OutputStream fileOut = new FileOutputStream("src/main/resources/quotation_word.docx")) {
            XWPFParagraph title = document.createParagraph();
            title.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun titleRun = title.createRun();
            titleRun.setBold(true);
            titleRun.setFontFamily("Arial");
            titleRun.setFontSize(20);
            titleRun.setText("Product Quote");

            addCompanyDetails(document);
            addCustomerDetails(document);
            addProductDetails(document, quoteData);
            addInfo(document);

            document.write(fileOut);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void addCompanyDetails(XWPFDocument document) {
        XWPFParagraph companyDetailsOne = document.createParagraph();
        XWPFRun companyDetailsOneRun = companyDetailsOne.createRun();

        for(int i = 0; i < 10; i++) {
            companyDetailsOneRun.addTab();
        }

        companyDetailsOneRun.setText("The Furniture Store");

        for(int i = 0; i < 11; i++) {
            companyDetailsOneRun.addTab();
        }
        try(InputStream is = new FileInputStream("src/main/resources/images/logo.jpg")) {
            companyDetailsOneRun.addPicture(is, XWPFDocument.PICTURE_TYPE_JPEG,
                    "logo.jpg", Units.pixelToEMU(140),
                    Units.pixelToEMU(64));

            XWPFParagraph companyDetailsTwo = document.createParagraph();
            XWPFRun companyDetailsTwoRun1 = companyDetailsTwo.createRun();
            companyDetailsTwoRun1.setText("Contact No: xxxxxxxxxx");

            for(int i = 0; i < 8; i++) {
                companyDetailsTwoRun1.addTab();
            }
            Date date = new Date();
            SimpleDateFormat format =  new SimpleDateFormat("dd/MM/yyyy");
            companyDetailsTwoRun1.setText("Date: " + format.format(date));

            companyDetailsTwoRun1.addBreak();
            XWPFRun companyDetailsTwoRun2 = companyDetailsTwo
                    .createHyperlinkRun("mailto:contact@furniture.com?subject=Inquiry");
            companyDetailsTwoRun2.setColor("0000FF");
            companyDetailsTwoRun2.setUnderline(UnderlinePatterns.SINGLE);
            companyDetailsTwoRun2.setText("Email: contact@furniture.com");

            companyDetailsTwoRun2.addBreak();

            XWPFRun companyDetailsTwoRun3 = companyDetailsTwo
                    .createHyperlinkRun("https://poi.apache.org/");
            companyDetailsTwoRun3.setColor("0000FF");
            companyDetailsTwoRun3.setUnderline(UnderlinePatterns.SINGLE);
            companyDetailsTwoRun3.setText("Web: www.furniture.com");

            companyDetailsTwoRun3.addBreak();

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
    }

    private void addCustomerDetails(XWPFDocument document) {
        XWPFParagraph customerDetails = document.createParagraph();
        XWPFRun customerDetailsRun = customerDetails.createRun();

        customerDetailsRun.setText("To:");
        customerDetailsRun.addBreak();
        customerDetailsRun.setText("Mr.John Doe");
        customerDetailsRun.addBreak();
        customerDetailsRun.setText("London");
        customerDetailsRun.addBreak();
        customerDetailsRun.setText("UK");
    }

    private void addProductDetails(XWPFDocument document, List<QuoteItem> quoteData) {
        XWPFTable productDetails = document.createTable();
        productDetails.setTableAlignment(TableRowAlign.CENTER);
        productDetails.setWidth(1440 * 6);
        XWPFTableRow firstRow = productDetails.getRow(0);
        XWPFTableCell qtyCell = firstRow.getCell(0);
        firstRow.getCell(0).removeParagraph(0);
        XWPFParagraph qtyCellPara = qtyCell.addParagraph();
        qtyCellPara.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun qtyCellRun = qtyCellPara.createRun();
        qtyCellRun.setText("Qty");
        qtyCellRun.setBold(true);
        qtyCellRun.setColor("0000FF");

        XWPFTableCell descriptionCell = firstRow.addNewTableCell();
        firstRow.getCell(1).removeParagraph(0);
        XWPFParagraph descriptionCellPara = descriptionCell.addParagraph();
        descriptionCellPara.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun descriptionCellRun = descriptionCellPara.createRun();
        descriptionCellRun.setText("Description");
        descriptionCellRun.setBold(true);
        descriptionCellRun.setColor("0000FF");

        XWPFTableCell priceCell = firstRow.addNewTableCell();
        firstRow.getCell(2).removeParagraph(0);
        XWPFParagraph priceCellPara = priceCell.addParagraph();
        priceCellPara.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun priceCellRun = priceCellPara.createRun();
        priceCellRun.setText("Unit Price");
        priceCellRun.setBold(true);
        priceCellRun.setColor("0000FF");
        priceCell.setWidthType(TableWidthType.DXA);
        priceCell.setWidth("1440");

        XWPFTableCell lineTotalCell = firstRow.addNewTableCell();
        firstRow.getCell(3).removeParagraph(0);
        XWPFParagraph lineTotalCellPara = lineTotalCell.addParagraph();
        lineTotalCellPara.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun lineTotalCellRun = lineTotalCellPara.createRun();
        lineTotalCellRun.setText("Line Total");
        lineTotalCellRun.setBold(true);
        lineTotalCellRun.setColor("0000FF");
        lineTotalCell.setWidthType(TableWidthType.DXA);
        lineTotalCell.setWidth("1440");

        Locale locale = new Locale("en", "US");
        NumberFormat currencyFormatter = NumberFormat.getCurrencyInstance(locale);

        double subTotal = 0;
        for(QuoteItem item : quoteData) {
            XWPFTableRow newRow = productDetails.createRow();
            XWPFTableCell cellOne = newRow.getCell(0);
            newRow.getCell(0).removeParagraph(0);
            XWPFParagraph cellOnePara = cellOne.addParagraph();
            cellOnePara.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun cellOneRun = cellOnePara.createRun();
            cellOneRun.setText(item.getQty() + "");

            XWPFTableCell cellTwo = newRow.getCell(1);
            newRow.getCell(1).removeParagraph(0);
            XWPFParagraph cellTwoPara = cellTwo.addParagraph();
            cellTwoPara.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun cellTwoRun = cellTwoPara.createRun();
            cellTwoRun.setText(item.getProduct().getName());

            XWPFTableCell cellThree = newRow.getCell(2);
            newRow.getCell(2).removeParagraph(0);
            XWPFParagraph cellThreePara = cellThree.addParagraph();
            cellThreePara.setAlignment(ParagraphAlignment.RIGHT);
            XWPFRun cellThreeRun = cellThreePara.createRun();
            cellThreeRun.setText(currencyFormatter.format(item.getProduct().getPrice()) + "");

            double lineTotal = item.getProduct().getPrice() * item.getQty();
            XWPFTableCell cellFour = newRow.getCell(3);
            newRow.getCell(3).removeParagraph(0);
            XWPFParagraph cellFourPara = cellFour.addParagraph();
            cellFourPara.setAlignment(ParagraphAlignment.RIGHT);
            XWPFRun cellFourRun = cellFourPara.createRun();
            cellFourRun.setText(currencyFormatter.format(lineTotal) + "");

            subTotal = subTotal + lineTotal;
        }

        XWPFTableRow subTotalRow = productDetails.createRow();

        XWPFTableCell  subTotalLblCell = subTotalRow.getCell(2);
        subTotalLblCell.removeParagraph(0);
        XWPFParagraph  subTotalLblCellPara = subTotalLblCell.addParagraph();
        XWPFRun subTotalLblCellRun = subTotalLblCellPara.createRun();
        subTotalLblCellRun.setBold(true);
        subTotalLblCellRun.setItalic(true);
        subTotalLblCellRun.setText("Sub Total");

        XWPFTableCell  subTotalValCell = subTotalRow.getCell(3);
        subTotalValCell.removeParagraph(0);
        XWPFParagraph  subTotalValCellPara = subTotalValCell.addParagraph();
        subTotalValCellPara.setAlignment(ParagraphAlignment.RIGHT);
        XWPFRun subTotalValCellRun = subTotalValCellPara.createRun();
        subTotalValCellRun.setText(currencyFormatter.format(subTotal) + "");

        XWPFTableRow taxRow = productDetails.createRow();
        XWPFTableCell  taxLblCell = taxRow.getCell(2);
        taxLblCell.removeParagraph(0);
        XWPFParagraph  taxLblCellPara = taxLblCell.addParagraph();
        XWPFRun taxLblCellRun = taxLblCellPara.createRun();
        taxLblCellRun.setBold(true);
        taxLblCellRun.setItalic(true);
        taxLblCellRun.setText("Tax");

        double tax = 25;
        XWPFTableCell  taxValCell = taxRow.getCell(3);
        taxValCell.removeParagraph(0);
        XWPFParagraph  taxValCellPara = taxValCell.addParagraph();
        taxValCellPara.setAlignment(ParagraphAlignment.RIGHT);
        XWPFRun taxValCellRun = taxValCellPara.createRun();
        taxValCellRun.setText(currencyFormatter.format(25) + "");

        XWPFTableRow totalRow = productDetails.createRow();
        XWPFTableCell  totalLblCell = totalRow.getCell(2);
        totalLblCell.removeParagraph(0);
        XWPFParagraph  totalLblCellPara = totalLblCell.addParagraph();
        XWPFRun totalLblCellRun = totalLblCellPara.createRun();
        totalLblCellRun.setBold(true);
        totalLblCellRun.setItalic(true);
        totalLblCellRun.setText("Total");

        double total = subTotal + tax;
        XWPFTableCell  totalValCell = totalRow.getCell(3);
        totalValCell.removeParagraph(0);
        XWPFParagraph  totalValCellPara = totalValCell.addParagraph();
        totalValCellPara.setAlignment(ParagraphAlignment.RIGHT);
        XWPFRun totalValCellRun = totalValCellPara.createRun();
        totalValCellRun.setText(currencyFormatter.format(total) + "");

    }

    private void addInfo(XWPFDocument document) {
        document.createParagraph();

        XWPFParagraph deliveryInfo = document.createParagraph();
        XWPFRun deliveryInfoRun = deliveryInfo.createRun();
        deliveryInfoRun.setText("Delivery can be arranged within city limits at a cost");

        document.createParagraph();
        XWPFParagraph generalInfo = document.createParagraph();
        generalInfo.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun generalInfoRun = generalInfo.createRun();
        generalInfoRun.setText("Should you have any inquiries please contact us");
    }

    public void extractText() {
        try(InputStream is =
                    new FileInputStream("src/main/resources/quotation_word.docx")) {

            XWPFDocument document   = new XWPFDocument(is);
            XWPFWordExtractor ext = new XWPFWordExtractor(document);
            System.out.println(ext.getText());
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void extractTable() {
        try (InputStream is = new FileInputStream("src/main/resources/quotation_word.docx")) {
            XWPFDocument document   = new XWPFDocument(is);
            List<XWPFTable> tables = document.getTables();
            if(tables != null && tables.size() > 0) {
                XWPFTable table = tables.get(0);
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        System.out.print(cell.getText() + " ");
                    }
                    System.out.println();
                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void extractImage() {
        try(InputStream is = new FileInputStream("src/main/resources/quotation_word.docx")) {
            XWPFDocument document   = new XWPFDocument(is);
            List<XWPFPictureData> pictures = document.getAllPictures();
            if(pictures != null && pictures.size() == 1) {
                XWPFPictureData picture = pictures.get(0);
                byte[] data = picture.getData();
                try(OutputStream out =
                            new FileOutputStream("src/main/resources/images/logo_ext_from_word.jpg")) {
                    out.write(data);
                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
