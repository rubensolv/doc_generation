package doc_generation;

import java.io.File;
import java.io.IOException;

import com.xandryex.WordReplacer;

public class MainClass {

	public static void main(String[] args) throws Exception {
		String partialPath = System.getProperty("user.dir").concat("/");
		
		File template = new File(partialPath+"Questionnaire2.docx");
		
		WordReplacer replacer = new WordReplacer(template);
		replacer.replaceWordsInTables("tag_uuid", args[0]);		
		replacer.replaceWordsInTables("tag_idade", args[1]);
		replacer.replaceWordsInTables("tag_sex", args[2]);
		replacer.replaceWordsInTables("tag_form", args[3]);
		replacer.replaceWordsInTables("tag_anos_xp", args[4]);
		replacer.replaceWordsInTables("tag_cursos", args[5]);
		replacer.replaceWordsInTables("tag_pesquisa", args[6]);
		replacer.replaceWordsInTables("tag_jogo_comp", args[7]);
		replacer.replaceWordsInTables("tag_joga_rts", args[8]);
		replacer.replaceWordsInTables("tag_sc1", args[9]);
		replacer.replaceWordsInTables("tag_sc2", args[10]);
		replacer.replaceWordsInTables("tag_sc3", args[11]);
		replacer.replaceWordsInTables("tag_sc4", args[12]);
		replacer.replaceWordsInTables("tag_sc5", args[13]);
		replacer.replaceWordsInTables("tag_sc6", args[14]);
		replacer.replaceWordsInTables("tag_sc7", args[15]);
		replacer.replaceWordsInTables("tag_sc8", args[16]);
		replacer.replaceWordsInTables("tag_sc9", args[17]);
		replacer.saveAndGetModdedFile(partialPath+"Questionnaire2_"+args[0]+".docx");

	}

}
