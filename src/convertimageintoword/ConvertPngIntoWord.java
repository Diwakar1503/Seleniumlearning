package convertimageintoword;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class ConvertPngIntoWord {

	public void getNoofImage(String filepath) throws FileNotFoundException, Exception {
		File file = new File(filepath);
		File[] fileList = file.listFiles();
		int sizeString = fileList.length;
		Scanner sc1 = new Scanner(System.in);
		System.out.println("Enter path to create Document");
		String word = sc1.nextLine();
		System.out.println("Enter Word document file name");
		String name = sc1.nextLine();
		sc1.close();
		String Name = word + "\\" + name + ".docx";
		XWPFDocument document = new XWPFDocument();
		FileOutputStream os = new FileOutputStream(new File(Name));
		XWPFParagraph paragraph = document.createParagraph();
		XWPFRun run = paragraph.createRun();
		try {
			for(File file1: fileList) {
				String image = file1.getAbsolutePath();
				int format=Document.PICTURE_TYPE_JPEG;
	             run.addBreak();
	             run.addPicture (new FileInputStream(image), format, image, Units.toEMU(460), Units.toEMU(250));
	             for(int i =0 ; i<6; i++) {
	                   run.addBreak();
	              }
			}
			
			document.write(os);
			os.close();
			document.close();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	public static void main(String[] args) throws Exception {

		ConvertPngIntoWord cpw = new ConvertPngIntoWord();
		Scanner sc = new Scanner(System.in);
		System.out.println("Enter the path contains Screenshot");
		String path = sc.nextLine();
		cpw.getNoofImage(path);
	}

}
