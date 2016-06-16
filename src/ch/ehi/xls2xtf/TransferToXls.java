package ch.ehi.xls2xtf;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import ch.ehi.basics.logging.EhiLogger;
import ch.ehi.basics.settings.Settings;
import ch.ehi.basics.view.GenericFileFilter;
import ch.interlis.ili2c.Ili2cException;
import ch.interlis.ili2c.Ili2cFailure;
import ch.interlis.ili2c.metamodel.AssociationDef;
import ch.interlis.ili2c.metamodel.AttributeDef;
import ch.interlis.ili2c.metamodel.Cardinality;
import ch.interlis.ili2c.metamodel.CompositionType;
import ch.interlis.ili2c.metamodel.EnumerationType;
import ch.interlis.ili2c.metamodel.LocalAttribute;
import ch.interlis.ili2c.metamodel.ObjectType;
import ch.interlis.ili2c.metamodel.RoleDef;
import ch.interlis.ili2c.metamodel.Table;
import ch.interlis.ili2c.metamodel.TransferDescription;
import ch.interlis.ili2c.metamodel.Type;
import ch.interlis.ili2c.metamodel.Viewable;
import ch.interlis.ili2c.metamodel.ViewableTransferElement;
import ch.interlis.iom.IomObject;
import ch.interlis.iom_j.Iom_jObject;
import ch.interlis.iom_j.itf.ModelUtilities;
import ch.interlis.iom_j.xtf.XtfWriter;
import ch.interlis.iox.IoxException;
import ch.interlis.iox_j.*;

public class TransferToXls 
{
public void doit(File xlsfile,File ilifile,Settings settings,String appHome) 
		throws FileNotFoundException, IOException, InvalidFormatException, DataException, IoxException
{
	
	HashMap<String,Viewable> classv=null;
	TransferDescription td=null;
	td=compileIli(ilifile,appHome,settings);
	if(td==null){
		return;
	}
	classv=ch.interlis.ili2c.generator.XSDGenerator.getTagMap(td);
	ch.interlis.ili2c.metamodel.Table classItem=(ch.interlis.ili2c.metamodel.Table) classv.get("CatalogueObjects_V1.Catalogues.Item");
	
	Workbook wb = null;
	if(GenericFileFilter.getFileExtension(xlsfile).equals("xls")){
		wb = new HSSFWorkbook();
	}else{
		wb = new XSSFWorkbook();
		
	}
	
	int tid=1;
	int bid=0;
	ch.interlis.ili2c.metamodel.Topic lastTopic=null;
	HashSet<String> sheetNames=new HashSet<String>();
	
	Iterator modeli=td.iterator();
	while(modeli.hasNext()){
		Object modelo=modeli.next();
		if(modelo instanceof ch.interlis.ili2c.metamodel.Model){
			ch.interlis.ili2c.metamodel.Model model=(ch.interlis.ili2c.metamodel.Model)modelo;
			if(model instanceof ch.interlis.ili2c.metamodel.PredefinedModel){
				continue;
			}
			if(model.getFileName().equals(ilifile.getPath())){
				//EhiLogger.debug("model.ilifile <"+model.getFileName()+">");
				Iterator topici=model.iterator();
				while(topici.hasNext()){
					Object topico=topici.next();
					if(topico instanceof ch.interlis.ili2c.metamodel.Topic){
						ch.interlis.ili2c.metamodel.Topic topic=(ch.interlis.ili2c.metamodel.Topic)topico;
						Iterator classi=topic.iterator();
						while(classi.hasNext()){
							Object classo=classi.next();
							if(classo instanceof ch.interlis.ili2c.metamodel.Table){
								ch.interlis.ili2c.metamodel.Table aclass=(ch.interlis.ili2c.metamodel.Table)classo;
								if(aclass.isExtending(classItem)){
									if(lastTopic!=topic){
										bid++;
									}
									EhiLogger.traceState("item "+aclass.getScopedName(null));
									String sheetName=aclass.getName();
									if(sheetNames.contains(sheetName)){
										sheetName=Integer.toString(tid)+"_"+sheetName;
									}
									sheetNames.add(sheetName);
									Sheet sheet = wb.createSheet(sheetName);
								    Row row = sheet.createRow((short)0);
								    row.createCell(0).setCellValue("BID");
								    row.createCell(1).setCellValue("CLASS");
								    row.createCell(2).setCellValue("TID");
								    ArrayList<String> colNames=getColNames(aclass);
								    {
									    int coli=3;
									    for(String colName:colNames){
										    row.createCell(coli++).setCellValue(colName);
									    }
								    }
								    // enumeration
								    if(true){
									    int coli=0;
						    			 java.util.ArrayList<Integer> enumColIdx=new java.util.ArrayList<Integer>();
						    			 java.util.ArrayList<java.util.ArrayList<String>> enumVals=new java.util.ArrayList<java.util.ArrayList<String>>();
									    for(String colName:colNames){
									    	AttributeDef attr = (AttributeDef) aclass.getElement(AttributeDef.class, colName);
									    	if(attr!=null){
									    		Type type=attr.getDomainResolvingAliases();
									    		if(type instanceof EnumerationType){
									    			 java.util.ArrayList<String> ev=new java.util.ArrayList<String>();
									    			 ModelUtilities.buildEnumList(ev,"",((EnumerationType) type).getConsolidatedEnumeration());
									    			 enumColIdx.add(coli+3);
									    			 enumVals.add(ev);
									    		}
									    	}
									    	coli++;
									    }
									    //for(int i=0;i<enumColIdx.size();i++){
									    int i=0;
										if(i<enumColIdx.size()){
									    	int enumCol=enumColIdx.get(i);
									    	java.util.ArrayList<String> ev=enumVals.get(i);
									    	for(int j=0;j<ev.size();j++){
									    		String enumVal=ev.get(j);
											    row = sheet.createRow((short)(j+1));
											    row.createCell(0).setCellValue(bid);
											    row.createCell(1).setCellValue(aclass.getScopedName(null));
											    row.createCell(2).setCellValue(tid);
											    row.createCell(enumCol).setCellValue(enumVal);
												tid++;
									    	}
									    }else{
										    row = sheet.createRow((short)1);
										    row.createCell(0).setCellValue(bid);
										    row.createCell(1).setCellValue(aclass.getScopedName(null));
										    row.createCell(2).setCellValue(tid);
											tid++;
									    }
								    }
									lastTopic=topic;
								}
							}
						}
					}
				}
			}else{
			}
		}
	}
	
	java.io.FileOutputStream fileout=new java.io.FileOutputStream(xlsfile);
	wb.write(fileout);
	fileout.close();
	wb.close();
}

private ArrayList<String> getColNames(Viewable v) {
	ArrayList<String> colNames=new ArrayList<String>();
	Iterator iter = v.getAttributesAndRoles2();
	while (iter.hasNext()) {
		ViewableTransferElement obj = (ViewableTransferElement)iter.next();
		if (obj.obj instanceof AttributeDef) {
			AttributeDef attr = (AttributeDef) obj.obj;
			if(!attr.isTransient()){
				Type proxyType=attr.getDomain();
				if(proxyType!=null && (proxyType instanceof ObjectType)){
					// skip implicit particles (base-viewables) of views
				}else{
					// map attr
					if(attr.getDomain() instanceof CompositionType){
						CompositionType type=(CompositionType) attr.getDomain();
						Table multiTextStruct=type.getComponentType();
						String structName=multiTextStruct.getScopedName(null);
						if(structName.equals("Localisation_V1.MultilingualMText")
							|| structName.equals("Localisation_V1.MultilingualText")
							|| structName.equals("LocalisationCH_V1.MultilingualMText")
							|| structName.equals("LocalisationCH_V1.MultilingualText")){
							LocalAttribute localTextAttr=(LocalAttribute) multiTextStruct.getElement(LocalAttribute.class,"LocalisedText");
							Table localTextStruct=((CompositionType)localTextAttr.getDomain()).getComponentType();
							IomObject multiText=null;
							for(String lang : new String[]{"de","fr","it","rm","en"}){
								colNames.add(attr.getName()+"["+lang+"]");
							}
						}
					}else{
						colNames.add(attr.getName());
					}
				}
			}
		}
		if(obj.obj instanceof RoleDef){
			RoleDef role = (RoleDef) obj.obj;
			// a role of an embedded association?
			if(obj.embedded){
				AssociationDef roleOwner = (AssociationDef) role.getContainer();
				if(roleOwner.getDerivedFrom()==null){
					// TODO map role
				}
			}
		}
	}
	return colNames;
}

private TransferDescription compileIli(File ilifile,String appHome,Settings settings) {
	ArrayList modeldirv=new ArrayList();
	String ilidirs=settings.getValue(Main.SETTING_ILIDIRS);
	EhiLogger.logState("ilidirs <"+ilidirs+">");
	String modeldirs[]=ilidirs.split(";");
	HashSet ilifiledirs=new HashSet();
	for(int modeli=0;modeli<modeldirs.length;modeli++){
		String m=modeldirs[modeli];
		if(m.equals(Main.ITF_DIR)){
			m=ilifile.getAbsoluteFile().getParentFile().getAbsolutePath();
			if(m!=null && m.length()>0){
				if(!modeldirv.contains(m)){
					modeldirv.add(m);				
				}
			}
		}else if(m.equals(Main.JAR_DIR)){
			m=appHome;
			if(m!=null){
				m=new java.io.File(m,"ilimodels").getAbsolutePath();
			}
			if(m!=null && m.length()>0){
				modeldirv.add(m);				
			}
		}else{
			if(m!=null && m.length()>0){
				modeldirv.add(m);				
			}
		}
	}		

	ch.interlis.ili2c.Main.setHttpProxySystemProperties(settings);
	TransferDescription td=null;
	ch.interlis.ili2c.config.Configuration ili2cConfig=null;
	ArrayList<String> requiredIliFiles=new ArrayList<String>();
	requiredIliFiles.add(ilifile.getPath());
	try {
		//ili2cConfig=ch.interlis.ili2c.ModelScan.getConfig(modeldirv, modelv);
		ch.interlis.ilirepository.IliManager modelManager=new ch.interlis.ilirepository.IliManager();
		modelManager.setRepositories((String[])modeldirv.toArray(new String[]{}));
		ili2cConfig=modelManager.getConfigWithFiles(requiredIliFiles);
		ili2cConfig.setGenerateWarnings(false);
	} catch (Ili2cException ex) {
		EhiLogger.logError(ex);
		return null;
	}

	try {
		ch.interlis.ili2c.Ili2c.logIliFiles(ili2cConfig);
		td=ch.interlis.ili2c.Ili2c.runCompiler(ili2cConfig);
	} catch (Ili2cFailure ex) {
		EhiLogger.logError(ex);
		return null;
	}
	return td;
}
}
