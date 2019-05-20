package TE.TransExtractor;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import org.w3c.dom.Document;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.NodeList;

import org.w3c.dom.Node;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import net.lingala.zip4j.core.ZipFile;
import net.lingala.zip4j.exception.ZipException;

/**
 * Hello world!
 *
 */
public class App {

	private static String patch = "\\xl\\drawings\\vmlDrawing1.vml";

	public static void main(String[] args) throws IOException, ZipException {

		File folder = new File("C:\\Users\\Rafal.Krakiewicz\\Desktop\\praca\\edi_translators\\");
		File[] listOfFiles = folder.listFiles();

		List<TranslatorBean> tbList = new ArrayList<TranslatorBean>();

		for (File file : listOfFiles) {
			if (file.isFile()) {
				System.out.println(file.getCanonicalPath());
				// System.out.println(file.getAbsolutePath());
				TranslatorBean tr = new TranslatorBean();
				getNames(file, tr);
				String unzipedFile = unzipXlsxFile(file);
				getTradosInfomation(unzipedFile, tr);
				String result = unzipXlsxFile(file);
				File results = new File(result);
				FileUtils.deleteDirectory(results);
				tbList.add(tr);
			}
		}

		PrintWriter writer = new PrintWriter("C:\\Users\\Rafal.Krakiewicz\\Desktop\\praca\\edi_translators\\log.txt",
				"UTF-8");

		for (TranslatorBean tr : tbList) {
			writer.println("('" + tr.getFirstName() + "','" + tr.getLastName() + "'," + tr.getSDL2014() + ","
					+ tr.getSDL2015() + "," + tr.getSDL2017() + "," + tr.getSDL2019() + ")");
		}

		writer.close();

	}

	private static TranslatorBean getTradosInfomation(String unzipedFile, TranslatorBean tr) {

		try {
			File fXmlFile = new File(unzipedFile + patch);
			DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
			DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
			Document doc = dBuilder.parse(fXmlFile);
			doc.getDocumentElement().normalize();

			NodeList nList = doc.getElementsByTagName("x:ClientData");
			printNote(nList, tr);

		} catch (Exception e) {
			System.out.println(e.getMessage());
		}

		return tr;
	}

	private static TranslatorBean getNames(File file, TranslatorBean tr) throws FileNotFoundException, IOException {
		File excelFile = file;
		FileInputStream fis = new FileInputStream(excelFile);
		XSSFWorkbook workbook = new XSSFWorkbook(fis);

		tr.setFirstName(workbook.getSheetAt(0).getRow(1).getCell(1).getStringCellValue());
		tr.setLastName(workbook.getSheetAt(0).getRow(2).getCell(1).getStringCellValue());
		tr.setSDL2014(2);
		tr.setSDL2015(2);
		tr.setSDL2017(2);
		tr.setSDL2019(2);

		workbook.close();
		fis.close();
		return tr;
	}

	public static String unzipXlsxFile(File file) throws IOException, ZipException {
		ZipFile zipFile = new ZipFile(file.getCanonicalPath());
		String unzipedFolder = FilenameUtils.getFullPath(file.getAbsolutePath())
				+ FilenameUtils.getBaseName(file.getName());
		zipFile.extractAll(unzipedFolder);
		return unzipedFolder;
	}

	private static void printNote(NodeList nodeList, TranslatorBean tr) {

		List<TradosBeam> tbList = createTBList(nodeList);

		setAnswer(tr, tbList);

		resetAnswers(tr);

		System.out.println("test");

	}

	private static void resetAnswers(TranslatorBean tr) {

		if (tr.getSDL2014() == 2) {
			tr.setSDL2014(0);
		}
		if (tr.getSDL2015() == 2) {
			tr.setSDL2015(0);
		}
		if (tr.getSDL2017() == 2) {
			tr.setSDL2017(0);
		}
		if (tr.getSDL2019() == 2) {
			tr.setSDL2019(0);
		}

	}

	private static void setAnswer(TranslatorBean tr, List<TradosBeam> tbList) {
		for (TradosBeam tb : tbList) {
			// System.out.println(tb.getCollumn() + " " + tb.getRow() + " " +
			// tb.getChecked());
			if (tb.getChecked() != null) {
				if (tb.getCollumn().equals("1") && tb.getRow().equals("4")) {
					if (tr.getSDL2014() == 0) {
						System.out.println("error");

					} else {
						tr.setSDL2014(1);
					}
				}
				if (tb.getCollumn().equals("1") && tb.getRow().equals("5")) {
					if (tr.getSDL2015() == 0) {
						System.out.println("error");

					} else {
						tr.setSDL2015(1);
					}
				}
				if (tb.getCollumn().equals("1") && tb.getRow().equals("6")) {
					if (tr.getSDL2017() == 0) {
						System.out.println("error");

					} else {
						tr.setSDL2017(1);
					}
				}
				if (tb.getCollumn().equals("1") && tb.getRow().equals("7")) {
					if (tr.getSDL2019() == 0) {
						System.out.println("error");

					} else {
						tr.setSDL2019(1);
					}
				}
				if (tb.getCollumn().equals("2") && tb.getRow().equals("4")) {
					if (tr.getSDL2014() == 1) {
						System.out.println("error");

					} else {
						tr.setSDL2014(0);
					}
				}
				if (tb.getCollumn().equals("2") && tb.getRow().equals("5")) {
					if (tr.getSDL2015() == 1) {
						System.out.println("error");

					} else {
						tr.setSDL2015(0);
					}
				}
				if (tb.getCollumn().equals("2") && tb.getRow().equals("6")) {
					if (tr.getSDL2017() == 1) {
						System.out.println("error");

					} else {
						tr.setSDL2017(0);
					}
				}
				if (tb.getCollumn().equals("2") && tb.getRow().equals("7")) {
					if (tr.getSDL2019() == 1) {
						System.out.println("error");

					} else {
						tr.setSDL2019(0);
					}
				}

			}
		}
	}

	private static List<TradosBeam> createTBList(NodeList nodeList) {

		List<TradosBeam> tbList = new ArrayList<TradosBeam>();

		for (int count = 0; count < nodeList.getLength(); count++) {

			TradosBeam tb = new TradosBeam();

			Node tempNode = nodeList.item(count);

			NodeList tempNodeChilds = tempNode.getChildNodes();

			for (int count2 = 0; count2 < tempNodeChilds.getLength(); count2++) {

				Node tempNodeChild = tempNodeChilds.item(count2);

				if (tempNodeChild.getNodeName().equals("x:Anchor")) {
					String anchorValue = tempNodeChild.getTextContent();
					getVaulesFromAnchor(anchorValue, tb);
				}

				if (tempNodeChild.getNodeName().equals("x:Checked")) {
					String checkedValue = tempNodeChild.getTextContent();
					tb.setChecked(checkedValue);
				}
			}

			tbList.add(tb);
		}

		return tbList;
	}

	private static TradosBeam getVaulesFromAnchor(String anchor, TradosBeam tb) {

		String[] tmp = anchor.split(",");

		tb.setCollumn(tmp[0].trim());
		tb.setRow(tmp[2].trim());

		return tb;
	}
}
