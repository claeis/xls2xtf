package ch.ehi.xls2xtf;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import ch.ehi.basics.logging.EhiLogger;
import ch.ehi.basics.settings.Settings;

public class Main {

		//TransferToXtf trsf = new TransferToXtf();
		//trsf.transfer("test/data/test1.xlsx");
	public static final String SETTING_ILIDIRS="ch.ehi.xls2xtf.ilidirs";
	public static final String SETTING_DIRUSED="ch.ehi.xls2xtf.dirused";
	public static final String ITF_DIR="%ITF_DIR";
	public static final String JAR_DIR="%JAR_DIR";
	/** name of application as shown to user.
	 */
	public static final String APP_NAME="xls2xtf";
	/** name of jar file.
	 */
	public static final String APP_JAR="xls2xtf.jar"; 
	private static String version=null;
	/** main program entry.
	 * @param args command line arguments.
	 */
	static public void main(String args[]){
		Settings settings=new Settings();
		settings.setValue(SETTING_ILIDIRS, ITF_DIR+";http://models.interlis.ch/;"+JAR_DIR);
		// arguments on export
		String xlsFile=null;
		String xtfFile=null;
		String httpProxyHost = null;
		String httpProxyPort = null;
		//if(args.length==0){
		readSettings(settings);
		//	ch.ehi.avgbs2txt.gui.MainFrame.main(avgbsFile,txtFile);
		//	return;
		//}
		int argi=0;
		boolean doGui=false;
		boolean initXls=false;
		for(;argi<args.length;argi++){
			String arg=args[argi];
			if(arg.equals("--trace")){
				EhiLogger.getInstance().setTraceFilter(false); 
			//}else if(arg.equals("--gui")){
			//	readSettings(settings);
			//	doGui=true;
			}else if(arg.equals("--initxls")){
				initXls=true;
				continue;
			}else if(arg.equals("--ilidirs")){
				argi++;
				settings.setValue(SETTING_ILIDIRS, args[argi]);
				continue;
			}else if(arg.equals("--proxy")) {
				    argi++;
				    settings.setValue(ch.interlis.ili2c.gui.UserSettings.HTTP_PROXY_HOST, args[argi]);
				    continue;
			}else if(arg.equals("--proxyPort")) {
				    argi++;
				    settings.setValue(ch.interlis.ili2c.gui.UserSettings.HTTP_PROXY_PORT, args[argi]);
				    continue;
			}else if(arg.equals("--version")){
				printVersion();
				return;
			}else if(arg.equals("--help")){
					printVersion ();
					System.err.println();
					printDescription ();
					System.err.println();
					printUsage ();
					System.err.println();
					System.err.println("OPTIONS");
					System.err.println();
					//System.err.println("--gui                 start GUI.");
					System.err.println("--ilidirs "+settings.getValue(SETTING_ILIDIRS)+" list of directories with ili-files.");
				    System.err.println("--proxy host          proxy server to access model repositories.");
				    System.err.println("--proxyPort port      proxy port to access model repositories.");
					System.err.println("--trace               enable trace messages.");
					System.err.println("--help                Display this help text.");
					System.err.println("--version             Display the version of "+APP_NAME+".");
					System.err.println();
					return;
				
			}else if(arg.startsWith("-")){
				EhiLogger.logAdaption(arg+": unknown option; ignored");
			}else{
				break;
			}
		}
		if(doGui){
			if(argi<args.length){
				xlsFile=args[argi];
				argi++;
			}
			if(argi<args.length){
				xtfFile=args[argi];
				argi++;
			}
			if(argi<args.length){
				EhiLogger.logAdaption(APP_NAME+": wrong number of arguments; ignored");
			}
			//ch.ehi.avgbs2txt.gui.MainFrame.main(avgbsFile,txtFile);
		}else{
			if(argi+2==args.length){
				xlsFile=args[argi];
				xtfFile=args[argi+1];
				if(initXls){
					runInitXls(xlsFile,xtfFile,settings);
				}else{
					runExport(xlsFile,xtfFile,settings);
				}
			}else{
				EhiLogger.logError(APP_NAME+": wrong number of arguments");
				return;
			}
		}
		
	}
	private final static String SETTINGS_FILE = System.getProperty("user.home") + "/.xls2xtf";
	public static void readSettings(Settings settings)
	{
		java.io.File file=new java.io.File(SETTINGS_FILE);
		try{
			if(file.exists()){
				settings.load(file);
			}
		}catch(java.io.IOException ex){
			EhiLogger.logError("failed to load settings from file "+SETTINGS_FILE,ex);
		}
	}
	public static void writeSettings(Settings settings)
	{
		java.io.File file=new java.io.File(SETTINGS_FILE);
		try{
			settings.store(file,APP_NAME+" settings");
		}catch(java.io.IOException ex){
			EhiLogger.logError("failed to settings settings to file "+SETTINGS_FILE,ex);
		}
	}
	/** main workhorse function.
	 * @param xlsFile name of xls file to be processed.
	 * @param xtfFile name of output file to be written.
	 */
	public static void runExport(
			String xlsFile,
			String xtfFile,
			Settings settings
		) {
		if(xlsFile==null  || xlsFile.length()==0){
			EhiLogger.logError("no XLS file given");
			return;
		}
		if(xtfFile==null  || xtfFile.length()==0){
			EhiLogger.logError("no XTF file given");
			return;
		}
		EhiLogger.logState(APP_NAME+"-"+getVersion());
		EhiLogger.logState("ili2c-"+ch.interlis.ili2c.Ili2c.getVersion());
		EhiLogger.logState("xlsFile <"+xlsFile+">");
		EhiLogger.logState("xtfFile <"+xtfFile+">");
		
		
		// process data file
		EhiLogger.logState("process data...");
		try{
			TransferToXtf trsf=new TransferToXtf();
			trsf.doit(new File(xlsFile),new File(xtfFile),settings,getAppHome());
			EhiLogger.logState("...conversion done");
		}catch(Throwable ex){
			EhiLogger.logError(ex);
		}
	}
	/** main workhorse function.
	 * @param xlsFile name of xls file to init.
	 * @param iliFile name of ili-model file to read/use.
	 */
	public static void runInitXls(
			String xlsFile,
			String iliFile,
			Settings settings
		) {
		if(xlsFile==null  || xlsFile.length()==0){
			EhiLogger.logError("no XLS file given");
			return;
		}
		if(iliFile==null  || iliFile.length()==0){
			EhiLogger.logError("no ILI file given");
			return;
		}
		EhiLogger.logState(APP_NAME+"-"+getVersion());
		EhiLogger.logState("ili2c-"+ch.interlis.ili2c.Ili2c.getVersion());
		EhiLogger.logState("xlsFile <"+xlsFile+">");
		EhiLogger.logState("iliFile <"+iliFile+">");
		
		
		// process data file
		EhiLogger.logState("process data...");
		try{
			TransferToXls trsf=new TransferToXls();
			trsf.doit(new File(xlsFile),new File(iliFile),settings,getAppHome());
			EhiLogger.logState("...conversion done");
		}catch(Throwable ex){
			EhiLogger.logError(ex);
		}
	}
	protected static void printVersion ()
	{
	  System.err.println(APP_NAME+" conversion, Version "+getVersion());
	  System.err.println("  Developed by Eisenhut Informatik AG, CH-3400 Burgdorf");
	}


	protected static void printDescription ()
	{
	  System.err.println("DESCRIPTION");
	  System.err.println("  Reads codelists from an MS-Excel file and converts it to an INTERLIS 2 transfer file.");
	}


	protected static void printUsage()
	{
	  System.err.println ("USAGE");
	  System.err.println("  java -jar xls2xtf.jar [Options] in.xls out.xtf");
	}
	/** get version of program.
	 * @return version e.g. "avgbs2txt-1.0.0"
	 */
	public static String getVersion() {
		  if(version==null){
		java.util.ResourceBundle resVersion = java.util.ResourceBundle.getBundle(ch.ehi.basics.i18n.ResourceBundle.class2qpackageName(Main.class)+".Version");
			// Major version numbers identify significant functional changes.
			// Minor version numbers identify smaller extensions to the functionality.
			// Micro versions are even finer grained versions.
			StringBuffer ret=new StringBuffer(20);
		ret.append(resVersion.getString("versionMajor"));
			ret.append('.');
		ret.append(resVersion.getString("versionMinor"));
			ret.append('.');
		ret.append(resVersion.getString("versionMicro"));
			ret.append('-');
		ret.append(resVersion.getString("versionDate"));
			version=ret.toString();
		  }
		  return version;
	}
	
	static public String getAppHome()
	{
	  String classpath = System.getProperty("java.class.path");
	  int index = classpath.toLowerCase().indexOf(APP_JAR);
	  int start = classpath.lastIndexOf(java.io.File.pathSeparator,index) + 1;
	  if(index > start)
	  {
		  return classpath.substring(start,index - 1);
	  }
	  return null;
	}

}
