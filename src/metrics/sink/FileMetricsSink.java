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
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class FileMetricsSink<ID> {

	protected MetricsTable<ID> table;

	public FileMetricsSink(String fileName) throws FileNotFoundException,
			InvalidFormatException, IOException {
		table = new MetricsTable<ID>(new File(fileName), Integer.MAX_VALUE, true);
	}

	public void flow(ID obj, String property, String value) {
		table.setProperty(obj, property, value);
	}

	public void flow(ID obj, String property, double value) {
		table.setProperty(obj, property, value);
	}

	public void terminate() throws IOException {
		table.dumpEntriesAndClose();
	}

}