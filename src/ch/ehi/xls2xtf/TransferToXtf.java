package ch.ehi.xls2xtf;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import ch.ehi.basics.logging.EhiLogger;
import ch.ehi.basics.settings.Settings;
import ch.interlis.ili2c.Ili2cException;
import ch.interlis.ili2c.Ili2cFailure;
import ch.interlis.ili2c.metamodel.AssociationDef;
import ch.interlis.ili2c.metamodel.AttributeDef;
import ch.interlis.ili2c.metamodel.Cardinality;
import ch.interlis.ili2c.metamodel.CompositionType;
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
import ch.interlis.iom_j.xtf.XtfWriter;
import ch.interlis.iox.IoxException;
import ch.interlis.iox_j.*;

public class TransferToXtf 
{
public void doit(File xlsfile,File xtffile,Settings settings,String appHome) 
		throws FileNotFoundException, IOException, InvalidFormatException, DataException, IoxException
{
	Workbook wb = WorkbookFactory.create(new java.io.FileInputStream(xlsfile));
	DataFormatter formatter = new DataFormatter();
	FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
	
	HashMap<String,Viewable> classv=null;
	HashSet<String> tids=new HashSet<String>();
	HashMap<String,ArrayList<IomObject>> baskets=new HashMap<String,ArrayList<IomObject>>();
	TransferDescription td=null;
	for (int k = 0; k < wb.getNumberOfSheets(); k++) {
		HashMap<String,Integer> colNames=null;
		int classCol=-1;
		int tidCol=-1;
		int bidCol=-1;
		Sheet sheet = wb.getSheetAt(k);
		int rows = sheet.getPhysicalNumberOfRows();
		EhiLogger.traceState("Sheet " + k + " \"" + wb.getSheetName(k) + "\" has " + rows+ " row(s)");
		for (int r = 0; r < rows; r++) {
			Row row = sheet.getRow(r);
			if (row == null) {
				continue;
			}

			int cells = row.getPhysicalNumberOfCells();
			int rowNum = row.getRowNum();
			EhiLogger.traceState("ROW " + rowNum + " has " + cells + " cell(s).");
			String recId="Row "+rowNum+": ";
			if(cells<3){
				EhiLogger.logError(recId+"incomplete row");
				continue;
			}
			if(colNames==null){
				ArrayList<String> cols=new ArrayList<String>();
				for (int c = 0; c < cells; c++) {
					Cell cell = row.getCell(c);

					String value = getCellStringValue(cell,formatter, evaluator);
					EhiLogger.traceState("CELL row="+rowNum+" , col=" + cell.getColumnIndex() + " VALUE=" + value);
					cols.add(value);
					if(colNames==null){
						if("CLASS".equals(value)){
							classCol=c;
						}else if("TID".equals(value)){
							tidCol=c;
						}else if("BID".equals(value)){
							bidCol=c;
						}
					}
				}
				if(classCol==-1){
					throw new DataException("missing column CLASS");
				}
				if(tidCol==-1){
					throw new DataException("missing column TID");
				}
				if(bidCol==-1){
					throw new DataException("missing column BID");
				}
				colNames=new HashMap<String,Integer>();
				int coli=0;
				for(String col : cols){
					if(coli>2){
						colNames.put(col, coli);
					}
					coli++;
				}
			}else{
				String bid=getCellStringValue(row.getCell(bidCol),formatter, evaluator);
				String className=getCellStringValue(row.getCell(classCol),formatter, evaluator);
				String tid=getCellStringValue(row.getCell(tidCol),formatter, evaluator);
				if(tids.contains(tid)){
					EhiLogger.logError(recId+"duplicate TID <"+tid+">");
				}else{
					tids.add(tid);
					if(td==null){
						td=compileIli(className,xtffile,appHome,settings);
						if(td==null){
							return;
						}
						classv=ch.interlis.ili2c.generator.XSDGenerator.getTagMap(td);
					}
					if(!classv.containsKey(className)){
						EhiLogger.logError(recId+"unknown class <"+className+">");
					}else{
						// map data
						Viewable aclass=classv.get(className);
						IomObject iomObj=new Iom_jObject(className,tid);
						aclass.getAttributesAndRoles2();
						mapObject(iomObj,aclass,row,colNames,formatter, evaluator);
						// save object
						ArrayList<IomObject> basket=null;
						if(!baskets.containsKey(bid)){
							basket=new ArrayList<IomObject>();
							baskets.put(bid, basket);
						}else{
							basket = baskets.get(bid);
						}
						basket.add(iomObj);
					}
				}
				
			}
		}
	}
	XtfWriter writer=new XtfWriter(xtffile,td);
	writer.write(new StartTransferEvent(Main.APP_NAME+"-"+Main.getVersion()));
	for(String bid:baskets.keySet()){
		ArrayList<IomObject> basket = baskets.get(bid);
		if(basket.size()>0){
			String names[]=basket.get(0).getobjecttag().split("\\.");
			String topic=names[0]+"."+names[1];
			writer.write(new StartBasketEvent(topic,bid));
			for(IomObject iomObj : basket){
				writer.write(new ObjectEvent(iomObj));
			}
			writer.write(new EndBasketEvent());
			
		}
	}
	writer.write(new EndTransferEvent());
	writer.close();
	writer=null;
}

private void mapObject(IomObject iomObj, Viewable v, Row row, HashMap<String, Integer> colNames, DataFormatter formatter, FormulaEvaluator evaluator) {
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
								if(colNames.containsKey(attr.getName()+"["+lang+"]")){
									Cell cell = row.getCell(colNames.get(attr.getName()+"["+lang+"]"));
									String value = getCellStringValue(cell,formatter, evaluator);
									if(value!=null){
										IomObject localText=new Iom_jObject(localTextStruct.getScopedName(null), null);
										localText.setattrvalue("Language", lang);
										localText.setattrvalue("Text", value);
										if(multiText==null){
											multiText=new Iom_jObject(structName, null);
										}
										multiText.addattrobj("LocalisedText", localText);
									}
								}
							}
							if(multiText!=null){
								iomObj.addattrobj(attr.getName(), multiText);
							}
						}
					}else{
						if(colNames.containsKey(attr.getName())){
							Cell cell = row.getCell(colNames.get(attr.getName()));
							String value = getCellStringValue(cell,formatter, evaluator);
							iomObj.setattrvalue(attr.getName(), value);
						}
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
	
}

private String getCellStringValue(Cell cell,DataFormatter formatter,
		FormulaEvaluator evaluator) {
	String value=null;
	if(cell==null){
		return null;
	}
	int cellType=cell.getCellType();
	if(cellType==Cell.CELL_TYPE_FORMULA){
		cellType=evaluator.evaluateFormulaCell(cell);					
	}
	switch (cellType) {

		case Cell.CELL_TYPE_FORMULA:
			value = cell.getCellFormula();
			break;

		case Cell.CELL_TYPE_NUMERIC:
			 if (DateUtil.isCellDateFormatted(cell)) {
	                //value=cell.getDateCellValue().toString();
	            	value=formatter.formatCellValue(cell);
	            } else {
	            	//String fmt=cell.getCellStyle().getDataFormatString();
	            	value=formatter.formatCellValue(cell);
	            	//value = Double.toString(cell.getNumericCellValue());
	            }
			break;

		case Cell.CELL_TYPE_STRING:
			value = cell.getRichStringCellValue().getString();
			break;
		default:
	}
	return value;
}

private TransferDescription compileIli(String aclass,File itffile,String appHome,Settings settings) {
	String names[]=aclass.split("\\.");
	String model=names[0];
	ArrayList modeldirv=new ArrayList();
	String ilidirs=settings.getValue(Main.SETTING_ILIDIRS);

	EhiLogger.logState("ilidirs <"+ilidirs+">");
	String modeldirs[]=ilidirs.split(";");
	HashSet ilifiledirs=new HashSet();
	for(int modeli=0;modeli<modeldirs.length;modeli++){
		String m=modeldirs[modeli];
		if(m.equals(Main.ITF_DIR)){
			m=itffile.getAbsoluteFile().getParentFile().getAbsolutePath();
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
	ArrayList<String> modelv=new ArrayList<String>();
	modelv.add(model);
	try {
		//ili2cConfig=ch.interlis.ili2c.ModelScan.getConfig(modeldirv, modelv);
		ch.interlis.ilirepository.IliManager modelManager=new ch.interlis.ilirepository.IliManager();
		modelManager.setRepositories((String[])modeldirv.toArray(new String[]{}));
		ili2cConfig=modelManager.getConfig(modelv, 0.0);
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
