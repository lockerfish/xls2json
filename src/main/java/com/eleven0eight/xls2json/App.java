/*
 * Copyright (c) 2012 Hendrix Tavarez
 *
 * Permission is hereby granted, free of charge, to any person obtaining
 * a copy of this software and associated documentation files (the
 * "Software"), to deal in the Software without restriction, including
 * without limitation the rights to use, copy, modify, merge, publish,
 * distribute, sublicense, and/or sell copies of the Software, and to
 * permit persons to whom the Software is furnished to do so, subject to
 * the following conditions:
 *
 * The above copyright notice and this permission notice shall be
 * included in all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
 * EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
 * MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
 * NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
 * LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
 * OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
 * WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 *
 */

package com.eleven0eight.xls2json;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.util.Iterator;
import java.util.ArrayList;

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

	public String convertXlsToJson(FileInputStream fis) throws Exception {

		Workbook workbook = WorkbookFactory.create(fis);
		Sheet sheet = workbook.getSheetAt(0);
		JSONObject json = new JSONObject();
		JSONArray items = new JSONArray();
		ArrayList cols = new ArrayList();

		for( int i=0; i <= sheet.getLastRowNum(); i++ ) {
			Row row = sheet.getRow(i);
			JSONObject item = new JSONObject();

			for(short colIndex=row.getFirstCellNum(); colIndex <= row.getLastCellNum(); colIndex++) {
				Cell cell = row.getCell(colIndex);
				if(cell == null) {
					continue;
				}
				if(i == 0) { // header
					cols.add( colIndex, cell.getStringCellValue() );
				} else {
					item.put((String)cols.get(colIndex), cell.getStringCellValue());
				}
			}
			if(item.length() > 0) {
				items.put(item);
			}
		}
		json.put("items", items);
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

	public String getJSON(String xlsFile) {

		try {
			FileInputStream fis = checkInputFile(xlsFile);
			if(fis != null) {
				return convertXlsToJson(fis);
			}
		} catch(Exception e) {
			// do nothing
		}

		return "";
	}

    public static void main( String[] args ) throws Exception {

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
				String json = app.convertXlsToJson(fis);
				System.out.println("saving json file " + outfile );
				app.saveJson(outfile, json);
			}
		}

    }
}
