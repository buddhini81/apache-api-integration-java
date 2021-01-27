package com.ourcompany.app;

import com.ourcompany.app.data.Product;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.*;

import java.awt.*;
import java.io.*;
import java.text.NumberFormat;
import java.util.List;
import java.util.Locale;

public class ProductPresentation {
    public void generateProductPresentation(List<Product> productData) {
        XMLSlideShow ppt = new XMLSlideShow();
        try(OutputStream os =
                    new FileOutputStream("src/main/resources/product_presentation.pptx")) {
            XSLFSlideMaster defaultMaster = ppt.getSlideMasters().get(0);
            XSLFSlideLayout titleLayout =
                    defaultMaster.getLayout(SlideLayout.TITLE);
            XSLFSlide firstSlide = ppt.createSlide(titleLayout);
            XSLFTextShape title = firstSlide.getPlaceholder(0);
            title.setText("Product Presentation");
            XSLFTextShape subTitle = firstSlide.getPlaceholder(1);
            subTitle.setText("The Furniture Store");

            addContentSlides(ppt, productData);

            XSLFSlide lastSlide = ppt.createSlide(titleLayout);
            XSLFTextShape thankyou = lastSlide.getPlaceholder(0);
            thankyou.setText("THANK YOU!");

            XSLFTextShape contact = lastSlide.getPlaceholder(1);
            contact.setText("The Furniture Store");
            XSLFTextRun contactRun = contact.addNewTextParagraph().addNewTextRun();
            contactRun.setFontSize(14d);
            contactRun.setText("Contact No: xxxxxxxxxx");

            XSLFTextParagraph emailPara = contact.addNewTextParagraph();
            XSLFTextRun emailRun = emailPara.addNewTextRun();
            emailRun.setFontSize(14d);
            emailRun.setText("Email: contact@furniture.com");
            XSLFHyperlink emailLink = emailRun.createHyperlink();
            emailLink.setAddress("mailto:contact@furniture.com?subject=Inquiry");

            XSLFTextParagraph webPara = contact.addNewTextParagraph();
            XSLFTextRun webRun = webPara.addNewTextRun();
            webRun.setFontSize(14d);
            webRun.setText("Web: www.furniture.com");
            XSLFHyperlink webLink = webRun.createHyperlink();
            webLink.setAddress("https://poi.apache.org/");

            ppt.write(os);

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void addContentSlides(XMLSlideShow ppt, List<Product> productData) {
        boolean catOneHeaderAdded = false;
        boolean catTwoHeaderOneAdded = false;
        boolean catTwoHeaderTwoAdded = false;

        for (Product p : productData) {
            if (p.getCategory() == 1 && !catOneHeaderAdded) {
                addSectionHeaderSlide(ppt, "FURNITURE", "Chairs");
                catOneHeaderAdded = true;
            } else if (p.getCategory() == 2 && p.getType() == 200 && !catTwoHeaderOneAdded) {
                addSectionHeaderSlide(ppt, "DECO", "Lamp Shades");
                catTwoHeaderOneAdded = true;
            } else if (p.getCategory() == 2 && p.getType() == 201 && !catTwoHeaderTwoAdded) {
                addSectionHeaderSlide(ppt, "DECO", "Wall Clocks");
                catTwoHeaderTwoAdded = true;
            }

            addProductSlide(ppt, p);
        }
    }

    private void addSectionHeaderSlide(XMLSlideShow ppt, String titleTxt, String subTitleTxt) {
        XSLFSlideLayout headerLayout =
                ppt.getSlideMasters().get(0).getLayout(SlideLayout.SECTION_HEADER);
        XSLFSlide slide = ppt.createSlide(headerLayout);
        XSLFTextShape title = slide.getPlaceholder(0);
        title.setText(titleTxt);
        XSLFTextShape subTitle = slide.getPlaceholder(1);
        subTitle.setText(subTitleTxt);
    }

    private void addProductSlide(XMLSlideShow ppt, Product product) {
        XSLFSlideLayout blankLayout =
                ppt.getSlideMasters().get(0).getLayout(SlideLayout.BLANK);
        XSLFSlide slide = ppt.createSlide(blankLayout);

        try(InputStream is = new FileInputStream(product.getImageFile())) {
            byte[] pictureData = IOUtils.toByteArray(is);
            XSLFPictureData pd =
                    ppt.addPicture(pictureData, XSLFPictureData.PictureType.JPEG);
            XSLFPictureShape pic = slide.createPicture(pd);
            pic.setAnchor(new Rectangle(100, 100, 225, 225));
            pic.setLineColor(Color.GRAY);
            pic.setLineWidth(5);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        XSLFTextBox nameText = slide.createTextBox();
        XSLFTextParagraph namePara = nameText.addNewTextParagraph();
        XSLFTextRun nameRun = namePara.addNewTextRun();
        nameRun.setText(product.getName());
        nameRun.setFontColor(Color.BLUE);
        nameRun.setFontSize(25d);
        namePara.addLineBreak();

        for (String s : product.getFeatures()) {
            XSLFTextParagraph featurePara = nameText.addNewTextParagraph();
            featurePara.setIndentLevel(1);
            featurePara.setBullet(true);
            //featurePara.setBulletCharacter("*");
            XSLFTextRun featureRun = featurePara.addNewTextRun();
            featureRun.setText(s);
            featureRun.setFontColor(Color.RED);
            featurePara.addLineBreak();
        }

        nameText.setAnchor(new Rectangle(350, 100, 400, 400));

        Locale locale = new Locale("en", "US");
        NumberFormat currencyFormatter = NumberFormat.getCurrencyInstance(locale);

        XSLFTextBox priceText = slide.createTextBox();
        XSLFTextParagraph pricePara = priceText.addNewTextParagraph();
        XSLFTextRun priceRun = pricePara.addNewTextRun();
        priceRun.setText(currencyFormatter.format(product.getPrice()) + "");
        priceRun.setFontFamily("Courier New");
        priceRun.setFontColor(Color.MAGENTA);
        priceRun.setBold(true);
        priceRun.setFontSize(34d);

        priceText.setAnchor(new Rectangle(100, 350,200,200));
    }

    public void reorderSlides() {
        try(InputStream is =
                    new FileInputStream("src/main/resources/product_presentation.pptx")) {
            XMLSlideShow ppt = new XMLSlideShow(is);
            XSLFSlide slideFour = ppt.getSlides().get(3);
            ppt.setSlideOrder(slideFour, 2);

            try(OutputStream os = new FileOutputStream("src/main/resources/product_presentation.pptx")) {
                ppt.write(os);
            }

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void removeLastSlide() {
        try(InputStream is =
                    new FileInputStream("src/main/resources/product_presentation.pptx")) {
            XMLSlideShow ppt = new XMLSlideShow(is);
            ppt.removeSlide(ppt.getSlides().size()-1);

            try(OutputStream os = new FileOutputStream("src/main/resources/product_presentation.pptx")) {
                ppt.write(os);
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
