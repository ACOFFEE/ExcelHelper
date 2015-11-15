package com.cf.excel.writer;

import java.io.File;

import org.apache.tools.ant.Project;
import org.apache.tools.ant.taskdefs.Zip;
import org.apache.tools.ant.types.FileSet;

public class ZipUtils {
	public static void zip(String srcPath, String outFile) {
		File srcdir = new File(srcPath);
		if (!srcdir.exists()) {
			return;
		}
		File zipFile = new File(outFile);
		Project prj = new Project();
		Zip zip = new Zip();
		zip.setProject(prj);
		zip.setDestFile(zipFile);
		FileSet fileSet = new FileSet();
		fileSet.setProject(prj);
		fileSet.setDir(srcdir);
		zip.addFileset(fileSet);
		zip.execute();
	}
}
