package data.xml;

import java.net.MalformedURLException;
import java.net.URI;
import java.net.URISyntaxException;
import java.net.URL;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.List;

import org.dom4j.Document;
import org.dom4j.DocumentHelper;
import org.dom4j.Element;

import data.excel.ExcelRow;

public class ConvertListToXMLDocument {

	public ConvertListToXMLDocument() {
	}

	public Document convert(List<ExcelRow> excelRowList) {
		Document document = DocumentHelper.createDocument();
		Element offers = setHeaderAndConstantData(document);
		for (int i = 0; i < excelRowList.size(); i++) {
			addItem(offers, excelRowList.get(i), i);
		}
		return document;
	}

	private void addItem(Element offers, ExcelRow item, int id) {
		Element offer = offers.addElement("offer");
		offer.addAttribute("id", Integer.toString(id));
		offer.addAttribute("type", "vendor.model");
		offer.addAttribute("available", item.getAvailable());
		offer.addAttribute("group_id", item.getGroup_id());
		// ------
		addElement(offer, "url", encodeURL(item.getUrl()));
		addElement(offer, "price", item.getPrice());
		addElement(offer, "currencyId", item.getCurrencyId());
		addElement(offer, "categoryId", item.getCategoryId());
		addElement(offer, "market_category", item.getMarket_category());
		addElement(offer, "picture", encodeURL(item.getPicture()));
		addElement(offer, "picture", encodeURL(item.getPicture2()));
		addElement(offer, "picture", encodeURL(item.getPicture3()));
		addElement(offer, "store", item.getStore());
		addElement(offer, "pickup", item.getPickup());
		addElement(offer, "delivery", item.getDelivery());
		addElement(offer, "local_delivery_cost", item.getLocal_delivery_cost());
		addElement(offer, "typePrefix", item.getTypePrefix());
		addElement(offer, "vendor", item.getVendor());
		addElement(offer, "model", item.getModel());
		addElement(offer, "description", item.getDescription());
		addElement(offer, "country_of_origin", item.getCountry_of_origin());
		// ---------------------------
		Element param = addElement(offer, "param", item.getYandex_color_name());
		addAttribute(param, "name", "Цвет");

		param = addElement(offer, "param", item.getSize());
		addAttribute(param, "name", "Размер");
		addAttribute(param, "unit", item.getSize_unit());

		param = addElement(offer, "param", item.getGender());
		addAttribute(param, "name", "Пол");

		param = addElement(offer, "param", item.getAge());
		addAttribute(param, "name", "Возраст");

		param = addElement(offer, "param", item.getMaterial());
		addAttribute(param, "name", "Материал");
	}

	private Element addAttribute(Element element, String name, String value) {
		if (element != null && name != null && !name.isEmpty() && value != null && !value.isEmpty()) {
			element.addAttribute(name, value);
		}
		return element;
	}

	private Element addElement(Element parentElement, String name, String value) {
		Element newElement = null;
		if (parentElement != null && name != null && !name.isEmpty() && value != null && !value.isEmpty()) {
			newElement = parentElement.addElement(name);
			newElement.setText(value);
		}
		return newElement;
	}

	private String encodeURL(String url) {
		String encodedURL = url;
		if (!url.isEmpty()) {
			try {
				URL urlObj = new URL(url);
				URI uri = new URI(urlObj.getProtocol(), urlObj.getUserInfo(), urlObj.getHost(), urlObj.getPort(), urlObj.getPath(),
						urlObj.getQuery(), urlObj.getRef());
				encodedURL = uri.toString();
			} catch (URISyntaxException | MalformedURLException e) {
				e.printStackTrace();
			}
		}
		return encodedURL;

	}

	private Element setHeaderAndConstantData(Document document) {
		document.addDocType("yml_catalog", null, "shops.dtd");
		Element yml_catalog = document.addElement("yml_catalog");
		yml_catalog.addAttribute("date", new SimpleDateFormat("yyyy-MM-dd HH:mm").format(Calendar.getInstance().getTime()));
		Element shop = yml_catalog.addElement("shop");
		shop.addElement("name").setText("Modnique");
		shop.addElement("company").setText("Modnique.com, Inc.");
		shop.addElement("url").setText("http://www.modnique.com");
		Element currencies = shop.addElement("currencies");
		currencies.addElement("currency").addAttribute("id", "RUR").addAttribute("rate", "1");
		currencies.addElement("currency").addAttribute("id", "USD").addAttribute("rate", "CBRF");

		Element categories = shop.addElement("categories");
		categories.addElement("category").addAttribute("id", "1").setText("Одежда");
		categories.addElement("category").addAttribute("id", "2").setText("Обувь");
		categories.addElement("category").addAttribute("id", "3").setText("Сумки");
		categories.addElement("category").addAttribute("id", "4").setText("Аксессуары");
		categories.addElement("category").addAttribute("id", "5").setText("Часы");
		categories.addElement("category").addAttribute("id", "6").setText("Очки");
		return shop.addElement("offers");
	}
}
