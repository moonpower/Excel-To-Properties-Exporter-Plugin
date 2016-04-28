package net.moon.excel.exporter.properties;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;
import org.eclipse.core.resources.IFile;
import org.eclipse.core.resources.IResource;
import org.eclipse.core.resources.IResourceVisitor;
import org.eclipse.core.runtime.CoreException;
import org.eclipse.core.runtime.IAdaptable;
import org.eclipse.core.runtime.Platform;
import org.eclipse.jface.viewers.ISelection;
import org.eclipse.jface.viewers.IStructuredSelection;
import org.eclipse.ui.handlers.HandlerUtil;

public class PropertiesFile extends AbstractHandler {

	@Override
	public Object execute(ExecutionEvent event) throws ExecutionException {
		ISelection sel = HandlerUtil.getActiveMenuSelection(event);
		IStructuredSelection selection = (IStructuredSelection) sel;
		Object[] objects = selection.toArray();
		for (Object obj : objects) {

			IResource resource = (IResource) Platform.getAdapterManager().getAdapter(obj, IResource.class);

			PropertiesFileCollector p = null;
			if (resource != null) {
				p = new PropertiesFileCollector();
				try {
					resource.accept(p);
				} catch (CoreException e) {
					e.printStackTrace();
				}
			}

			for (Object name : p.info.keySet().toArray()) {
				String infoKey = (String) name;
				ArrayList<IFile> list = p.info.get(infoKey);
				WritableWorkbook workbook = null;
				WritableSheet createSheet = null;
				int currentRow = 0;
				int currentColumn = 0;
				Map<String, Properties> propertiesMap = new HashMap<String, Properties>();
				for (int i = 0; i < list.size(); i++) {
					IFile f = (IFile) list.get(i);
					String projectName = f.getProject().getName();
					String path = f.getFullPath().toPortableString();
					String directoryPath = path.substring(path.indexOf(f.getFullPath().segment(1)), path.length()).replace("/", ".");
					String fileName = projectName + "@" + directoryPath + ".xls";

					try {
						Properties properties = new Properties();
						properties.load(f.getContents());
						if (f.getName().contains("ko")) {
							propertiesMap.put("ko", properties);
						} else if (f.getName().contains("ja")) {
							propertiesMap.put("ja", properties);
						} else {
							propertiesMap.put("default", properties);
						}
						if (workbook == null) {
							File createFile = new File("c:/test/" + fileName);
							workbook = Workbook.createWorkbook(createFile);
							createSheet = workbook.createSheet("sheet1", 0);
							createSheet.addCell(new Label(currentColumn, currentRow, "NAME"));
							currentColumn += 1;
							createSheet.addCell(new Label(currentColumn, currentRow, "DEFAULT"));
							currentColumn += 1;
							createSheet.addCell(new Label(currentColumn, currentRow, "KO"));
							currentColumn += 1;
							createSheet.addCell(new Label(currentColumn, currentRow, "JA"));
							currentRow += 1;
							currentColumn = 0;
							
						}

					} catch (FileNotFoundException e) {
						e.printStackTrace();
					} catch (IOException e) {
						e.printStackTrace();
					} catch (CoreException e) {
						e.printStackTrace();
					} catch (RowsExceededException e) {
						e.printStackTrace();
					} catch (WriteException e) {
						e.printStackTrace();
					}
				}
				Properties defaultProperties = propertiesMap.get("default");
				Properties koProperties = propertiesMap.get("ko");
				Properties jaProperties = propertiesMap.get("ja");
				if (defaultProperties != null) {
					for (Object each : defaultProperties.keySet().toArray()) {
						String key = (String) each;
						try {
							createSheet.addCell(new Label(currentColumn, currentRow, key));
							currentColumn += 1;
							createSheet.addCell(new Label(currentColumn, currentRow, (String) defaultProperties.get(key)));
							if (koProperties != null && koProperties.get(key) != null) {
								currentColumn += 1;
								createSheet.addCell(new Label(currentColumn, currentRow, (String) koProperties.get(key)));
							}
							if (jaProperties != null && jaProperties.get(key) != null) {
								currentColumn += 1;
								createSheet.addCell(new Label(currentColumn, currentRow, (String) jaProperties.get(key)));
							}
							currentRow += 1;
							currentColumn = 0;
							
						} catch (RowsExceededException e) {
							e.printStackTrace();
						} catch (WriteException e) {
							e.printStackTrace();
						}
					}
				}

				createSheet.setColumnView(0, 20);
				createSheet.setColumnView(1, 40);
				createSheet.setColumnView(2, 40);
				createSheet.setColumnView(3, 40);
				
				try {
					workbook.write();
					workbook.close();
				} catch (IOException e) {
					e.printStackTrace();
				} catch (WriteException e) {
					e.printStackTrace();
				}
			}
		}
		return null;
	}
}
