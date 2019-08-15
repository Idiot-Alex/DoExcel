package com.hotstrip.excel.usermodel;


import com.hotstrip.annotation.Column;
import com.hotstrip.enums.ExcelTypeEnums;
import com.hotstrip.exception.DoExcelException;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.*;


/**
 * 导出工具类
 * @author Lius
 */
public class Export {
	private static Logger logger = LoggerFactory.getLogger(Export.class);


	/**
	 * 导出方法
	 * @param title      sheet 标题
	 * @param pojoClass  实体对象
	 * @param dataSet    实体集合
	 * @param out        输出
	 * @param type    xlsx   xls
	 */
	private static Workbook writeWorkBook(String title, Class pojoClass, Collection dataSet, OutputStream out, ExcelTypeEnums type) throws IOException, NoSuchMethodException, InvocationTargetException, IllegalAccessException {
		// 首先检查数据看是否是正确的
		if (dataSet == null || dataSet.size() == 0) {
			throw new DoExcelException("导出数据为空！");
		}
		if (title == null || out == null || pojoClass == null) {
			throw new DoExcelException("导出数据，传入参数不能为空！");
		}
		// 声明一个工作薄
		Workbook workbook = null;
		if(type.equals(ExcelTypeEnums.XLSX))  //  .xlsx
			workbook = new XSSFWorkbook();
		else		//	.xls
			workbook = new HSSFWorkbook();
		// 生成一个表格
		Sheet sheet = workbook.createSheet(title);

		// 标题
		List<String> exportFieldTitle = new ArrayList<String>();
		List<Integer> exportFieldWidth = new ArrayList<Integer>();
		// 拿到所有列名，以及导出的字段的get方法
		List<Method> methodObj = new ArrayList<Method>();
		Map<String, Method> convertMethod = new HashMap<String, Method>();
		// 得到所有字段
		Field fileds[] = pojoClass.getDeclaredFields();
		//属性的simpleType
		String filedTypes[] = new String[fileds.length]; //长度就是上面获取字段属性的个数
		// 遍历整个filed
		for (int i = 0; i < fileds.length; i++) {
			Field field = fileds[i];
			//判断该属性是否含有指定注解
			if(field.isAnnotationPresent(Column.class)){
				if(!"serialVersionUID".equals(field.getName()))         //排除序列化id
					filedTypes[i] = field.getType().getSimpleName();	//获取属性类型
			}
			Column doSheet = field.getAnnotation(Column.class);
			// 如果设置了annottion
			if (doSheet != null) {
				// 添加到标题
				exportFieldTitle.add(doSheet.name());
				// 添加标题的列宽
				exportFieldWidth.add(doSheet.width());
				// 添加到需要导出的字段的方法
				String fieldname = field.getName();
				// System.out.println(i+"列宽"+excel.exportName()+" "+excel.exportFieldWidth());
				StringBuffer getMethodName = new StringBuffer("get");
				getMethodName.append(fieldname.substring(0, 1)
						.toUpperCase());
				getMethodName.append(fieldname.substring(1));

				Method getMethod = pojoClass.getMethod(
						getMethodName.toString(), new Class[] {});

				methodObj.add(getMethod);
				/*if (excel.exportConvertSign() == 1) {
					StringBuffer getConvertMethodName = new StringBuffer("get");
					getConvertMethodName.append(fieldname.substring(0, 1).toUpperCase());
					getConvertMethodName.append(fieldname.substring(1));
					getConvertMethodName.append("Convert");
					Method getConvertMethod = pojoClass.getMethod(getConvertMethodName.toString(),new Class[] {});
					convertMethod.put(getMethodName.toString(),getConvertMethod);
				}*/
			}
		}
		int index = 0;
		// 产生表格标题行
		Row row = sheet.createRow(index);
		for (int i = 0, exportFieldTitleSize = exportFieldTitle.size(); i < exportFieldTitleSize; i++) {
			Cell cell = row.createCell(i);
			// cell.setCellStyle(style);
			RichTextString text = null;
			if(type.equals(ExcelTypeEnums.XLSX))
				text = new XSSFRichTextString(exportFieldTitle.get(i));
			else
				text = new HSSFRichTextString(exportFieldTitle.get(i));
			cell.setCellValue(text);
		}

		//设置单元格的样式
		CellStyle cellstyle = getCellStyle(workbook);

		// 设置每行的列宽
		for (int i = 0; i < exportFieldWidth.size(); i++) {
			// 256=65280/255
			sheet.setColumnWidth(i, 256 * exportFieldWidth.get(i));
		}
		Iterator its = dataSet.iterator();
		// 循环插入剩下的集合
		while (its.hasNext()) {
			// 从第二行开始写，第一行是标题
			index++;
			row = sheet.createRow(index);
			Object t = its.next();
			for (int k = 0, methodObjSize = methodObj.size(); k < methodObjSize; k++) {
				Cell cell = row.createCell(k);
				//设置样式
				cell.setCellStyle(cellstyle);
				Method getMethod = methodObj.get(k);
				Object value = null;
				if (convertMethod.containsKey(getMethod.getName())) {
					Method cm = convertMethod.get(getMethod.getName());
					value = cm.invoke(t, new Object[] {});
				} else {
					value = getMethod.invoke(t, new Object[] {});
				}
				// 对null属性赋值
				if (value == null) {
					cell.setCellValue("");
				} else {
					cell.setCellValue(String.valueOf(value));
				}
			}
		}
		workbook.write(out);
		return workbook;
	}

	private static CellStyle getCellStyle(Workbook workBook) {
		//创建一个样式
		CellStyle cellStyle = workBook.createCellStyle();
		//创建一个DataFormat对象
		DataFormat format = workBook.createDataFormat();
		//这样才能真正的控制单元格格式，@就是指文本型
		cellStyle.setDataFormat(format.getFormat("@"));
		return cellStyle;
	}

	public static void main(String[] args) throws Exception {
		/*// 构造一个模拟的List来测试，实际使用时，这个集合用从数据库中查出来
		Student pojo2 = new Student();
		pojo2.setName("第一行数据");
		pojo2.setAge(28);
		pojo2.setSex(2);
		pojo2.setBirthDate(ZonedDateTime.now().toLocalDateTime());
		pojo2.setDesc("abcdefghijklmnop");
		pojo2.setIsVip(true);
		List<Student> list = new ArrayList<Student>();
		list.add(pojo2);
		for (int i = 0; i < 50000; i++) {
			Student pojo = new Student();
			pojo.setName("一二三四五六七八九");
			pojo.setAge(22);
			pojo.setSex(1);
			pojo.setDesc("abcdefghijklmnop");
			pojo.setBirthDate(ZonedDateTime.now().toLocalDateTime());
			pojo.setIsVip(false);
			list.add(pojo);
		}
		// 构造输出对象，可以从response输出，直接向用户提供下载
		OutputStream out = new FileOutputStream("D://testOne.xls");
		//OutputStream out1=response.getOutputStream();

		// 开始时间
		Long l = System.currentTimeMillis();
		// 注意
		ExcelExport.exportExcel("测试", Student.class, list, out);
		out.close();
		// 结束时间
		Long s = System.currentTimeMillis();
		System.out.println("excel导出成功");
		System.out.println("总共耗时：" + (s - l));*/
	}

}
