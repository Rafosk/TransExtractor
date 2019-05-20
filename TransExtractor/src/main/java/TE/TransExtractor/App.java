package TE.TransExtractor;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import org.w3c.dom.Document;

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

		File file = new File(
				"C:\\Users\\Rafal.Krakiewicz\\Desktop\\praca\\edi_translators\\SDL Trados Update_Elisa Colombo.xlsx");

		TranslatorBean tr = new TranslatorBean();
		getNames(file,tr);
		String unzipedFile = unzipXlsxFile(file);
		getTradosInfomation(unzipedFile,tr);
		String result = unzipXlsxFile(file);
		System.out.println(result);

	}

	private static TranslatorBean getTradosInfomation(String unzipedFile, TranslatorBean tr) {

		try {

			File fXmlFile = new File(unzipedFile);
			DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
			DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
			Document doc = dBuilder.parse(fXmlFile);
			doc.getDocumentElement().normalize();

			System.out.println("Root element :" + doc.getDocumentElement().getNodeName());
		} catch (Exception e) {
			// TODO: handle exception
		}
		
		
		
		return tr;
	}

	private static TranslatorBean getNames(File file, TranslatorBean tr) throws FileNotFoundException, IOException {
		File excelFile = file;		
		FileInputStream fis = new FileInputStream(excelFile);
		XSSFWorkbook workbook = new XSSFWorkbook(fis);

		tr.setFirstName(workbook.getSheetAt(0).getRow(1).getCell(1).getStringCellValue());
		tr.setLastName(workbook.getSheetAt(0).getRow(2).getCell(1).getStringCellValue());

		workbook.close();
		fis.close();
		return tr;
	}

	public static String unzipXlsxFile(File file) throws IOException, ZipException {
		ZipFile zipFile = new ZipFile(file.getCanonicalPath());				
		String unzipedFolder = FilenameUtils.getFullPath(file.getAbsolutePath())
				+ FilenameUtils.getBaseName(file.getName());
		System.out.println(unzipedFolder);
		zipFile.extractAll(unzipedFolder);
		return unzipedFolder;

	}
}
