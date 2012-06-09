package com.eleven0eight.xls2json;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.util.Iterator;

import org.json.JSONObject;
import org.json.JSONArray;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;

/**
 * App to convert XLS files to JSON
 *
 */
public class App {

	public String convertXls2Json(FileInputStream fis) throws Exception {

//		FileInputStream fis = new FileInputStream("/home/hendrix/Desktop/test1.xls");

		Workbook workbook = WorkbookFactory.create(fis);

		Sheet sheet = workbook.getSheetAt(0);

		JSONObject json = new JSONObject();

		JSONArray rows = new JSONArray();

		for( Iterator<Row> rowsIT = sheet.rowIterator(); rowsIT.hasNext(); ) {

			Row row = rowsIT.next();
			JSONObject jRow = new JSONObject();

			JSONArray cells = new JSONArray();

			for( Iterator<Cell> cellsIT = row.cellIterator(); cellsIT.hasNext(); ) {
				Cell cell = cellsIT.next();
				cells.put( cell.getStringCellValue() );
			}
			jRow.put( "cell", cells );
			rows.put ( jRow );
		}

		json.put( "rows", rows );

		return json.toString();

	}

	public FileInputStream checkInputFile(String filename) throws Exception {

		File file = new File(filename);

		if(file.exists()) {
			return new FileInputStream(filename);
		}

		System.err.println("ERROR: " + filename + " does exists.");
		return null;
	}

	public void saveJson(String filename, String json) throws Exception {

		BufferedWriter out = new BufferedWriter(new FileWriter(filename));
		out.write(json);
		out.close();

	}

    public static void main( String[] args ) throws Exception {

		// read args
		// check if args is valid
		// do convertion
		// save json file

		if( args == null || args.length < 2) {
			System.err.println("ERROR: input and/or output files are missing.");

			System.out.println("\t USAGE:");
			System.out.println("\t  java -cp target/xls2json-1.0-jar-with-dependencies.jar com.eleven0eight.xls2json.App {inputfile} {outputfile}");
		} else {
			String filename = args[0];
			String outfile = args[1];
			App app = new App();

			System.out.println("checking if file " + filename  + " exists.");
			FileInputStream fis = app.checkInputFile(filename);
			if( fis != null) {
				System.out.println("converting file " + filename  + " to JSON.");
				String json = app.convertXls2Json(fis);
				System.out.println("saving json file " + outfile );
				app.saveJson(outfile, json);
			}
		}

    }
}
