package com.abc.word2Html.util;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.PicturesManager;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.usermodel.PictureType;
import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.core.FileURIResolver;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.w3c.dom.Document;

public class Word2Html {
	public static void main(String[] args) throws Throwable {
		String path = "E:\\";
		String file = "考前资料-管理.doc";
		toHtml(path,file);

	}
	
	private static void toHtml(String path,String file) throws Throwable{
		File f = new File(path+file);
		if(!f.exists()){
			System.out.println("文件不存在!!!");
			return;
		}
		if(file.endsWith("doc")){
			doc(f);
		}else if(file.endsWith("docx")){
			docx(f);
		}
	}
	
	private static String getFileName(File f){
		String fileName = f.getName();
		String result = fileName.substring(0,fileName.indexOf(".docx")==-1?fileName.indexOf(".doc"):fileName.indexOf(".docx"));
		return result;
		// CRC32 c = new CRC32();
		// c.update(result.getBytes());
		// return c.getValue();
	}
	
	private static void docx(File f) throws Throwable {

		// 生成 XWPFDocument
		InputStream in = new FileInputStream(f);
		XWPFDocument document = new XWPFDocument(in);

		// 准备 XHTML 选项 (设置 IURIResolver，把图片放到文件绝对路径下image/word/media文件夹
		File imageFolderFile = new File(f.getParentFile().getPath()+"image\\"+getFileName(f));
		XHTMLOptions options = XHTMLOptions.create().URIResolver(new FileURIResolver(imageFolderFile));
		options.setExtractor(new FileImageExtractor(imageFolderFile));
		options.setIgnoreStylesIfUnused(false);
		options.setFragment(true);

		// 将XWPFDocument 转换为  XHTML
		File file1 = new File(f.getParentFile().getPath()+getFileName(f)+".html");
		OutputStream out = new FileOutputStream(file1);
		XHTMLConverter.getInstance().convert(document, out, options);
	}

	public static  void doc(final File f) throws Throwable{
		// 生成 HWPFDocument 
		InputStream input = new FileInputStream(f);
		HWPFDocument wordDocument = new HWPFDocument(input);
		// 把图片放到文件绝对路径下image/word/media文件夹
		WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(
				DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());
		wordToHtmlConverter.setPicturesManager(new PicturesManager() {
			public String savePicture(byte[] content, PictureType pictureType, String suggestedName, float widthInches,
					float heightInches) {
				// 与docx保持一致
				File imgFile = new File(f.getParentFile().getPath()+"image\\word\\media\\");
				if(!imgFile.exists()){
					imgFile.mkdirs();
				}
				try {
					FileOutputStream out = new FileOutputStream(imgFile.getPath()+"\\"+suggestedName);
					out.write(content);
					out.close();
				} catch (Exception e) {
					e.printStackTrace();
				}
				
				return "image/word/media/"+suggestedName;
			}
		});
		wordToHtmlConverter.processDocument(wordDocument);
		Document htmlDocument = wordToHtmlConverter.getDocument();
		ByteArrayOutputStream outStream = new ByteArrayOutputStream();
		DOMSource domSource = new DOMSource(htmlDocument);
		StreamResult streamResult = new StreamResult(outStream);
		TransformerFactory tf = TransformerFactory.newInstance();
		Transformer serializer = tf.newTransformer();
		serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
		serializer.setOutputProperty(OutputKeys.INDENT, "yes");
		serializer.setOutputProperty(OutputKeys.METHOD, "html");
		serializer.transform(domSource, streamResult);
		outStream.close();
		String content = new String(outStream.toByteArray());
		FileUtils.writeStringToFile(new File(f.getParentFile().getPath(), getFileName(f)+".html"), content, "utf-8");
	}
}