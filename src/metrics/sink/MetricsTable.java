/*
 * This is free software; you can redistribute it and/or modify it
 * under the terms of the GNU Lesser General Public License as
 * published by the Free Software Foundation; either version 2.1 of
 * the License, or (at your option) any later version.
 *
 * This software is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this software; if not, write to the Free
 * Software Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA
 * 02110-1301 USA, or see the FSF site: http://www.fsf.org.
 */

package metrics.sink;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Collection;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Set;
import java.util.TreeSet;

import org.apache.commons.collections.MapIterator;
import org.apache.commons.collections.bidimap.TreeBidiMap;
import org.apache.commons.collections.keyvalue.DefaultKeyValue;
import org.apache.commons.collections.map.MultiValueMap;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;

public class MetricsTable<ID> {

	/*
	 * When an element reaches MAX_NUMBER_PROPERTIES number of properties, the
	 * entry is automatically dumped into the file, allowing the reference to be
	 * removed from the map.
	 */
	private final int MAX_NUMBER_PROPERTIES;

	/*
	 * Maps an element to to a collection of key-value pairs. These key-value
	 * pairs represent the properties of the entry.
	 */
	private MultiValueMap map = MultiValueMap.decorate(new LinkedHashMap());

	/*
	 * Maps int -> String, where int is the column position and the String is
	 * the column (property) name
	 */
	private TreeBidiMap columnMapping = new TreeBidiMap();

	private int columnCount = 1;
	private int rowCount = 0;

	/*
	 * Created on the File passed as argument to the ctor.
	 */
	private Workbook workBook;

	/*
	 * Sheet created on the WorkBook.
	 */
	private Sheet currentSheet;

	/*
	 * Flag to keep track if the columns headers cells were created.
	 */
	private boolean headersWerePrint = false;

	/*
	 * File in which the workBook will the written to.
	 */
	private final File output;

	public MetricsTable(File output, int max, boolean preserveExistingWorkbook)
			throws FileNotFoundException, IOException, InvalidFormatException {
		if (max < 1)
			throw new IllegalArgumentException(
					"Maximum number of properties must be greater than 0. ");
		MAX_NUMBER_PROPERTIES = max;

		if (output == null)
			throw new IllegalArgumentException(
					"Argument file must not be null. ");
		this.output = output;
		init(output, preserveExistingWorkbook);
	}

	private void init(File output, boolean preserveExistingWorkbook)
			throws InvalidFormatException, FileNotFoundException, IOException {
		if (!output.exists())
			output.createNewFile();
		if (preserveExistingWorkbook)
			try {
				workBook = WorkbookFactory.create(new FileInputStream(output));
			} catch (IllegalArgumentException ex) {

			}
		if (workBook == null)
			workBook = new HSSFWorkbook();
		this.currentSheet = workBook.createSheet();
	}

	public void setProperty(ID id, String property, String value) {
		if (id == null)
			throw new IllegalArgumentException("The element must not be null. ");
		map.put(id, new DefaultKeyValue(property, value));
		if (map.size(id) == MAX_NUMBER_PROPERTIES) {
			dumpEntry(id, map.getCollection(id));
			map.remove(id);
		}
		if (!columnMapping.containsValue(property)) {
			columnMapping.put(columnCount, property);
			columnCount++;
		}
	}

	public void setProperty(ID id, String property, Number value) {
		map.put(id, new DefaultKeyValue(property, value));
		if (map.size(id) == MAX_NUMBER_PROPERTIES) {
			dumpEntry(id, map.getCollection(id));
			map.remove(id);
		}
		if (!columnMapping.containsValue(property)) {
			columnMapping.put(columnCount, property);
			columnCount++;
		}
	}

	private void dumpEntry(ID id, Collection<DefaultKeyValue> properties) {
		if (!headersWerePrint) {
			printHeaders();
			headersWerePrint = true;
		}
		Row entryRow = currentSheet.createRow(rowCount++);

		Cell methodSignatureCell = entryRow.createCell(0);
		methodSignatureCell.setCellValue(id.toString());

		Iterator<DefaultKeyValue> iterator = properties.iterator();
		while (iterator.hasNext()) {
			DefaultKeyValue nextKeyVal = iterator.next();

			String property = (String) nextKeyVal.getKey();
			Object value = nextKeyVal.getValue();
			Integer columnIndex = (Integer) columnMapping.getKey(property);

			if (value instanceof Number) {
				Cell cell = entryRow.createCell(columnIndex);
				cell.setCellValue((Double) value);
				cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);

			} else {
				Cell cell = entryRow.createCell(columnIndex);
				cell.setCellValue((String) value);
				cell.setCellType(HSSFCell.CELL_TYPE_STRING);
			}
		}
	}

	private void printHeaders() {
		MapIterator columnMapIterator = columnMapping.mapIterator();
		Row headerRow = currentSheet.createRow(0);
		rowCount++;

		while (columnMapIterator.hasNext()) {
			Cell headerCell = headerRow.createCell((Integer) columnMapIterator
					.next());
			headerCell.setCellValue((String) columnMapIterator.getValue());
		}
	}

	private void printFooters() {
		int columns = columnMapping.size();
		Row firstRow = currentSheet.getRow(1);
		Row lastRow = currentSheet.getRow(rowCount - 1);

		Row sumFooterRow = currentSheet.createRow(rowCount++);
		Cell sumFooterLabelCell = sumFooterRow.createCell(0);
		sumFooterLabelCell.setCellValue("SUM");

		Row averageFooterRow = currentSheet.createRow(rowCount++);
		Cell averageFooterLabelCell = averageFooterRow.createCell(0);
		averageFooterLabelCell.setCellValue("AVERAGE");

		for (int index = 1; index <= columns; index++) {
			Cell cell = firstRow.getCell(index);
			if (cell == null) {
				cell = firstRow.createCell(index);
			}
			Cell sumFooterCell = sumFooterRow.createCell(index);
			Cell averageFooterCell = averageFooterRow.createCell(index);

			CellReference firstCell = new CellReference(firstRow.getCell(index));
			Cell lastRowCell = lastRow.getCell(index);
			if (lastRowCell == null) {
				lastRowCell = lastRow.createCell(index);
			}
			CellReference lastCell = new CellReference(lastRowCell);

			sumFooterCell.setCellFormula("SUM(" + firstCell.formatAsString()
					+ ":" + lastCell.formatAsString() + ")");
			averageFooterCell.setCellFormula("AVERAGE("
					+ firstCell.formatAsString() + ":"
					+ lastCell.formatAsString() + ")");
		}
	}

	public void dumpEntriesAndClose() throws IOException {
		dumpAllEntries();
		printFooters();
		int noOfColumns = columnMapping.keySet().size();
		for (int index = 0; index < noOfColumns; index++) {
			currentSheet.autoSizeColumn(index);
		}
		FileOutputStream outStream = new FileOutputStream(output);
		workBook.write(outStream);
		outStream.close();
	}

	private void dumpAllEntries() {
		Set keySet = map.keySet();
		for (Object object : keySet) {
			dumpEntry((ID) object, map.getCollection(object));
		}
		map.clear();
	}
}