package com.ourcompany.app;

import com.ourcompany.app.data.Product;
import com.ourcompany.app.data.QuoteItem;

import java.util.ArrayList;
import java.util.List;

/**
 * Hello world!
 *
 */
public class App {
    public static void main( String[] args ) {
        List<Product> productData = init();
//        ProductCatalog productCatalog = new ProductCatalog();
//        productCatalog.generateCatalog(productData);
//        productCatalog.loadAndModifyProductCatalog();
//        productCatalog.extractContent();
//        //productCatalog.removePage();
//        //productCatalog.setXMPMetadata();
//        productCatalog.setImageMetaData();
//        productCatalog.readImageMetaData();
//        productCatalog.setDocumentInformation();
//        List<Product> moreProductData = loadMoreProducts();
//        productCatalog.addPage(moreProductData.subList(0,3), "APPLIANCES - LAMP SHADES");
//        productCatalog.addPage(moreProductData.subList(3,6), "APPLIANCES - WALL CLOCKS");
//        productCatalog.splitPDF();
//        productCatalog.mergePDFs();
//        productCatalog.encryptPDF();

//        Quotation quotation = new Quotation();
        //     List<QuoteItem> quoteData = loadQuoteData();
//        quotation.generateQuotation(quoteData, false);
//
//        quotation.readQuoteDate();
//        quotation.readLogoImage();
//        quotation.readHyperLink();

        //WordQuotation wordQuotation = new WordQuotation();
        //wordQuotation.generateQuotation(quoteData);

        //wordQuotation.extractText();
        //wordQuotation.extractTable();
        //wordQuotation.extractImage();

        List<Product> allProducts =  loadAllProducts();
        ProductPresentation presentation = new ProductPresentation();
        presentation.generateProductPresentation(allProducts);
        presentation.reorderSlides();
        presentation.removeLastSlide();
    }

    private static List<Product> init() {
        List<Product> list = new ArrayList<>();
        Product prd1 = new Product();
        prd1.setCategory(1);
        prd1.setName("Craftatoz Wooden Chair");
        prd1.setFeatures(List.of("100 % Sheesham wood", "Finest polishing", "Creative curves and contours"));
        prd1.setPrice(399);
        prd1.setImageFile("src/main/resources/images/Craftatoz_Wooden_Chair.jpg");
        list.add(prd1);

        Product prd2 = new Product();
        prd2.setCategory(1);
        prd2.setName("Premium Rome Series Wood Chair");
        prd2.setFeatures(List.of("Teak Wood", "Polished", "Protective Foot Glide Insert"));
        prd2.setPrice(299);
        prd2.setImageFile("src/main/resources/images/Premium_Rome_Series_Wood_Chair.jpg");
        list.add(prd2);

        Product prd3 = new Product();
        prd3.setCategory(1);
        prd3.setName("Wish Chair - Beech / Natural");
        prd3.setFeatures(List.of("Beech Wood Frame", "Paper Rope Seat", "Steam Bent Back Rail"));
        prd3.setPrice(299);
        prd3.setImageFile("src/main/resources/images/Wish Chair_Beech_Natural.jpg");
        list.add(prd3);

        return list;
    }

    private static List<Product> loadMoreProducts() { // Add a new set of products
        List<Product> list = new ArrayList<>();
        Product prd1 = new Product();
        prd1.setCategory(2);
        prd1.setType(200);
        prd1.setName("Natural Rattan Accent Lamp Shade");
        prd1.setFeatures(List.of("Crafted of natural rattan", "10.2\"Dia. x 6.7\"H", "Works with uno socket lamp bases only"));
        prd1.setPrice(17);
        prd1.setImageFile("src/main/resources/images/Natural_Rattan_Accent_Lamp Shade.jpg");
        list.add(prd1);

        Product prd2 = new Product();
        prd2.setCategory(2);
        prd2.setType(200);
        prd2.setName("Beige Lamp Shade");
        prd2.setFeatures(List.of("2 pack transitional hardback", "7\" Dia. x 8.6\"H"));
        prd2.setPrice(20);
        prd2.setImageFile("src/main/resources/images/Beige_Lamp Shade.jpg");
        list.add(prd2);

        Product prd3 = new Product();
        prd3.setCategory(2);
        prd3.setType(200);
        prd3.setName("Brass Pierced Table Lamp Shade");
        prd3.setFeatures(List.of("Handcrafted of iron with antique brass finish", "Works with uno socket lamp bases only"));
        prd3.setPrice(50);
        prd3.setImageFile("src/main/resources/images/Brass Pierced_Table Lamp Shade.jpg");
        list.add(prd3);

        Product prd4 = new Product();
        prd4.setCategory(2);
        prd4.setType(201);
        prd4.setName("Gabriel Clock");
        prd4.setFeatures(List.of("High-quality silent quartz movement", "Industry-leading 5-year manufacturer's warranty"));
        prd4.setPrice(109);
        prd4.setImageFile("src/main/resources/images/Gabriel_Clock.jpg");
        list.add(prd4);

        Product prd5 = new Product();
        prd5.setCategory(2);
        prd5.setType(201);
        prd5.setName("Fenton Wall Clock");
        prd5.setFeatures(List.of("Time Display: Analog", "Shape: Round","Material: Plastic"));
        prd5.setPrice(37);
        prd5.setImageFile("src/main/resources/images/Fenton_Wall_Clock.jpg");
        list.add(prd5);

        Product prd6 = new Product();
        prd6.setCategory(2);
        prd6.setType(201);
        prd6.setName("Wooden Metal Wall Clock");
        prd6.setFeatures(List.of("Classical black metal hands", "large roman numerals","Silent non-ticking"));
        prd6.setPrice(106);
        prd6.setImageFile("src/main/resources/images/Wooden_Metal_Wall_Clock.jpg");
        list.add(prd6);

        return list;
    }

    private static List<QuoteItem> loadQuoteData() {
        List<QuoteItem> list = new ArrayList<>();
        QuoteItem item1 = new QuoteItem();
        item1.setQty(6);
        Product prd1 = new Product();
        prd1.setName("Craftatoz Wooden Chair");
        prd1.setPrice(399);
        item1.setProduct(prd1);
        list.add(item1);

        QuoteItem item2 = new QuoteItem();
        item2.setQty(1);
        Product prd2 = new Product();
        prd2.setName("Wish Chair - Beech / Natural");
        prd2.setPrice(299);
        item2.setProduct(prd2);
        list.add(item2);

        return list;
    }

    private static List<Product> loadAllProducts() {
        List<Product> products = init();
        List<Product> moreProducts = loadMoreProducts();
        products.addAll(moreProducts);

        return products;
    }
}
