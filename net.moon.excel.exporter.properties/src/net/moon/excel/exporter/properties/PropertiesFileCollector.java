package net.moon.excel.exporter.properties;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

import org.eclipse.core.resources.IFile;
import org.eclipse.core.resources.IResource;
import org.eclipse.core.resources.IResourceVisitor;
import org.eclipse.core.runtime.CoreException;

public class PropertiesFileCollector implements IResourceVisitor {

	public ArrayList<IFile> result;

	public Map<Object,ArrayList<IFile>> info;
	
	public PropertiesFileCollector(){
		this.result = new ArrayList<IFile>();
		this.info = new HashMap<Object,ArrayList<IFile>>();
	}
	
	@Override
	public boolean visit(IResource resource) throws CoreException {
		if (resource instanceof IFile) {
			IFile file = (IFile) resource;
			String ext = file.getProjectRelativePath().getFileExtension();
			if (ext != null && ext.toLowerCase().equals("properties")) { //$NON-NLS-1$
				String portableString = file.getParent().getFullPath().toPortableString();
				ArrayList<IFile> value = info.get(portableString);
				if(value == null){
					info.put(portableString, new ArrayList());
				}
				info.get(portableString).add(file);
//				this.result.add(file);
			}
			
			return false;
		}
		return true;
	}

}
