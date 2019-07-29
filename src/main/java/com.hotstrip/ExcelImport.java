package com.hotstrip;



import com.hotstrip.annotation.DoSheet;
import com.hotstrip.enums.ExcelTypeEnums;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.Type;
import java.util.*;

/**
 * 导入工具类
 * @author Lius
 */
public class ExcelImport {
	private static Logger logger = LoggerFactory.getLogger(ExcelImport.class);

	public static Collection importExcel(File file, Class pojoClass, ExcelTypeEnums type, String... patternExcelTypeEnums) {

		Collection dist = new ArrayList();
		try {
			// 得到目标目标类的所有的字段列表
			Field filed[] = pojoClass.getDeclaredFields();
			// 将所有标有Annotation的字段，也就是允许导入数据的字段,放入到一个map中
			Map<String, Method> fieldSetMap = new HashMap<String, Method>();

			Map<String, Method> fieldSetConvertMap = new HashMap<String, Method>();

			// 循环读取所有字段
			for (int i = 0; i < filed.length; i++) {
				Field f = filed[i];

				// 得到单个字段上的Annotation
				DoSheet doSheet = f.getAnnotation(DoSheet.class);

				// 如果标识了Annotationd的话
				if (doSheet != null) {
					// 构造设置了Annotation的字段的Setter方法
					String fieldname = f.getName();
					String setMethodName = "set"
							+ fieldname.substring(0, 1).toUpperCase()
							+ fieldname.substring(1);
					// 构造调用的method，
					Method setMethod = pojoClass.getMethod(setMethodName,
							new Class[] { f.getType() });
					// 将这个method以Annotaion的名字为key来存入。

					// 对于重名将导致 覆盖失败，对于此处的限制需要
					fieldSetMap.put(doSheet.name(), setMethod);
					/*if (doSheet.importConvertSign() == 1) {
						StringBuffer setConvertMethodName = new StringBuffer("set");
						setConvertMethodName.append(fieldname.substring(0, 1).toUpperCase());
						setConvertMethodName.append(fieldname.substring(1));
						setConvertMethodName.append("Convert");
						Method getConvertMethod = pojoClass.getMethod(
								setConvertMethodName.toString(),
								new Class[] { String.class });
						// System.out.println(excel.exportName()+"--"+setConvertMethodName.toString());
						fieldSetConvertMap.put(doSheet.name(), getConvertMethod);
					}*/

				}
			}

			// 将传入的File构造为FileInputStream;
			FileInputStream in = new FileInputStream(file);
			// // 得到工作表
			Workbook workbook = null;
			if(type.equals(ExcelTypeEnums.XLSX))  //  .xlsx
				workbook = new XSSFWorkbook(in);
			else		//	.xls
				workbook = new HSSFWorkbook(in);
			// // 得到第一页
			Sheet sheet = workbook.getSheetAt(0);
			// // 得到第一面的所有行
			Iterator<Row> row = sheet.rowIterator();

			// 得到第一行，也就是标题行
			Row title = row.next();
			// 得到第一行的所有列
			Iterator<Cell> cellTitle = title.cellIterator();
			// 将标题的文字内容放入到一个map中。
			Map titlemap = new HashMap();
			// 从标题第一列开始
			int i = 0;
			// 循环标题所有的列
			while (cellTitle.hasNext()) {
				Cell cell = cellTitle.next();
				String value = cell.getStringCellValue();
				titlemap.put(i, value);
				i = i + 1;
			}

			while (row.hasNext()) {
				// 标题下的第一行
				Row rown = row.next();
				
				// 行的所有列
				Iterator<Cell> cellbody = rown.cellIterator();
				// 得到传入类的实例
				Object tObject = pojoClass.newInstance();

				int k = 0;
				// 遍历一行的列
				while (cellbody.hasNext()) {
					Cell cell = cellbody.next();
					if(cell.getCellType() == CellType.NUMERIC){ //是number类型
						cell.setCellType(CellType.STRING);
					}
					// 这里得到此列的对应的标题
					String titleString = (String) titlemap.get(k);
					// 如果这一列的标题和类中的某一列的Annotation相同，那么则调用此类的的set方法，进行设值
					if (fieldSetMap.containsKey(titleString)) {
						Method setMethod = (Method) fieldSetMap.get(titleString);
						// 得到setter方法的参数
						Type[] ts = setMethod.getGenericParameterTypes();
						// 只要一个参数
						String xclass = ts[0].toString();
						
						// 判断参数类型

						if (fieldSetConvertMap.containsKey(titleString)) {
							fieldSetConvertMap.get(titleString).invoke(tObject,
									cell.getStringCellValue());
						} else {
							if (xclass.equals("class java.lang.String")) {
								if(cell.getCellType() == CellType.NUMERIC){ //是number类型
									cell.setCellType(CellType.STRING);
								}
								setMethod.invoke(tObject, cell.getStringCellValue());
							} else if (xclass.equals("class java.util.Date")) {
								setMethod.invoke(tObject, cell.getDateCellValue());
							} else if (xclass.equals("class java.lang.Boolean")) {
								setMethod.invoke(tObject, cell.getBooleanCellValue());
							} else if (xclass.equals("class java.lang.Integer")) {
								setMethod.invoke(tObject, new Integer(cell.getStringCellValue()));
							} else if (xclass.equals("class java.lang.Long")) {
								setMethod.invoke(tObject, new Long(cell.getStringCellValue()));
							}
						}
					}
					// 下一列
					k = k + 1;
				}
				dist.add(tObject);
			}
		} catch (Exception e) {
			logger.error(e.getMessage(), e);
			return null;
		}
		return dist;
	}

}
