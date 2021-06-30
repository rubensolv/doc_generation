package doc_generation;

import java.io.File;
import java.io.IOException;

import com.xandryex.WordReplacer;

import java.util.HashMap;
import java.util.StringTokenizer;

import org.docx4j.XmlUtils;
import org.docx4j.jaxb.Context;
import org.docx4j.model.datastorage.migration.VariablePrepare;
import org.docx4j.openpackaging.io.SaveToZipFile;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;

public class MainClass {

	public static void main(String[] args) throws Exception {
		String partialPath = System.getProperty("user.dir").concat("/");
		
		File template = new File(partialPath+"Questionnaire2.docx");
		
		// Exclude context init from timing
		org.docx4j.wml.ObjectFactory foo = Context.getWmlObjectFactory();

		// Input docx has variables in it: ${colour}, ${icecream}
		String inputfilepath = System.getProperty("user.dir") + "/Questionnaire2.docx";

		boolean save = true;
		String outputfilepath = System.getProperty("user.dir") + "/Questionnaire2_"+args[0]+".docx";

		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new java.io.File(inputfilepath));
		VariablePrepare.prepare(wordMLPackage);		
		MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();

		HashMap<String, String> mappings = new HashMap<String, String>();
		mappings.put("tag_uuid", args[0]);
		mappings.put("tag_idade", args[1]);
		mappings.put("tag_sex", args[2]);
		mappings.put("tag_form", args[3]);
		mappings.put("tag_anos_xp", args[4]);
		mappings.put("tag_cursos", args[5]);
		mappings.put("tag_pesquisa", args[6]);
		mappings.put("tag_jogo_comp", args[7]);
		mappings.put("tag_joga_rts", args[8]);		
		mappings.put("tag_sc1", newtab(newlineToBreakHack(args[9])));
		mappings.put("tag_sc2", newtab(newlineToBreakHack(args[10])));
		mappings.put("tag_sc3", newtab(newlineToBreakHack(args[11])));
		mappings.put("tag_sc4", newtab(newlineToBreakHack(args[12])));
		mappings.put("tag_sc5", newtab(newlineToBreakHack(args[13])));
		mappings.put("tag_sc6", newtab(newlineToBreakHack(args[14])));
		mappings.put("tag_sc7", newtab(newlineToBreakHack(args[15])));
		mappings.put("tag_sc8", newtab(newlineToBreakHack(args[16])));
		mappings.put("tag_sc9", newtab(newlineToBreakHack(args[17])));

		long start = System.currentTimeMillis();

		// Approach 1 (from 3.0.0; faster if you haven't yet caused unmarshalling to
		// occur):

		documentPart.variableReplace(mappings);

		/*
		 * // Approach 2 (original)
		 * 
		 * // unmarshallFromTemplate requires string input String xml =
		 * XmlUtils.marshaltoString(documentPart.getJaxbElement(), true); // Do it...
		 * Object obj = XmlUtils.unmarshallFromTemplate(xml, mappings); // Inject result
		 * into docx documentPart.setJaxbElement((Document) obj);
		 */

		long end = System.currentTimeMillis();
		long total = end - start;
		System.out.println("Time: " + total);

		// Save it
		if (save) {
			SaveToZipFile saver = new SaveToZipFile(wordMLPackage);
			saver.save(outputfilepath);
		} else {
			System.out.println(XmlUtils.marshaltoString(documentPart.getJaxbElement(), true, true));
		}
		

	}
	
	
	/**
	 * Hack to convert a new line character into w:br.
	 * If you need this sort of thing, consider using 
	 * OpenDoPE content control data binding instead.
	 *  
	 * @param r
	 * @return
	 */
	private static String newlineToBreakHack(String r) {
		r = newDoubleTab(r);
//		StringTokenizer st = new StringTokenizer(r, "\\&"); // tokenize on the newline character, the carriage-return character, and the form-feed character
//		StringBuilder sb = new StringBuilder();
//		
//		boolean firsttoken = true;
//		while (st.hasMoreTokens()) {						
//			String line = (String) st.nextToken();
//			if (firsttoken) {
//				firsttoken = false;
//			} else {
//				sb.append("</w:t><w:br/><w:t>");				
//			}
//			sb.append(line);
//		}
//		return sb.toString();
		return r.replace("&", "</w:t><w:br/><w:t>");
	}
	
	private static String newtab(String r) {
		
//		StringTokenizer st = new StringTokenizer(r, "\\#"); // tokenize on the newline character, the carriage-return character, and the form-feed character
//		StringBuilder sb = new StringBuilder();
//		
//		boolean firsttoken = true;
//		while (st.hasMoreTokens()) {						
//			String line = (String) st.nextToken();
//			if (firsttoken) {
//				firsttoken = false;
//			} else {				
//				sb.append("</w:t><w:tab/><w:t>");
//			}
//			sb.append(line);
//		}
//		return sb.toString();	
		return r.replaceAll("#", "</w:t><w:tab/><w:t>");
	}
	
	private static String newDoubleTab(String r) {

//		StringTokenizer st = new StringTokenizer(r, "\\#\\#"); // tokenize on the newline character, the carriage-return character, and the form-feed character
//		StringBuilder sb = new StringBuilder();
//		
//		boolean firsttoken = true;
//		while (st.hasMoreTokens()) {						
//			String line = (String) st.nextToken();
//			if (firsttoken) {
//				firsttoken = false;
//			} else {				
//				sb.append("</w:t><w:tab/><w:t>");
//				sb.append(" </w:t><w:tab/><w:t>");
//			}
//			sb.append(line);
//		}
//		return sb.toString();
		
		return r.replaceAll("##", "</w:t><w:tab/><w:t></w:t><w:tab/><w:t>");
	}

}
