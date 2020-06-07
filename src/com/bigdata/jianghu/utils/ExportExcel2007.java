package com.bigdata.jianghu.utils;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 生成excel 2007版
 * 
 * @author lilingchen
 * @param <T>
 *
 *            @ version 1.0v
 */
public class ExportExcel2007 {
	private String outPath;
	private XSSFWorkbook wb;
	private OutputStream out;

	public ExportExcel2007() {

	}

	public ExportExcel2007(String outPath, XSSFWorkbook wb) {
		this.outPath = outPath;
		this.wb = wb;
		this.init();
	}

	public void init() {
		try {
			out = new FileOutputStream(outPath);
		} catch (FileNotFoundException e) {
			System.out.println("excel输出流创建异常");
			e.printStackTrace();
		}
	}

	/**
	 * 按 语义指纹 id 排重
	 * 
	 * @param sheet
	 * @param dateList
	 * @param headers
	 * @param columns
	 * @param idmap
	 * @param fingmap
	 */
	Map<Long, Long> idmap = new HashMap<Long, Long>();
	Map<Long, Long> fingmap = new HashMap<Long, Long>();
	Map<Long, Long> tfingmap = new HashMap<Long, Long>();

	public void createExcelDistinctNoSort(XSSFSheet sheet,
                                          List<Map<String, String>> dateList, List<String> headers,
                                          List<String> columns) {
		XSSFRow row;
		if (headers != null || headers.size() > 0) {
			row = sheet.createRow((int) 0);
			for (int i = 0; i < headers.size(); i++) {
				String titleName = StringUtils.trimToEmpty(headers.get(i));
				if (!"".equals(titleName)&&!titleName.equals("titlewithurl")) {
					Cell cell = row.createCell((short) i);
					cell.setCellValue(titleName);
					// CellStyle style = wb.createCellStyle(); //设置单元格颜色
					// style.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
					// style.setFillPattern(CellStyle.SOLID_FOREGROUND);
					// cell.setCellStyle(style);
				}
			}
		}
		int rowCount = sheet.getLastRowNum();
		if (dateList != null && dateList.size() > 0) {

			int i = 0;// 行数
			for (Map map : dateList) {
				// j为列数
				row = sheet.createRow((int) rowCount + i + 1);
//				long idkey = Long.parseLong((String) map.get(columns.get(columns.indexOf("id"))));
//				String fingStr = (String) map.get(columns[1]);
//				String tfingStr = (String) map.get(columns[2]);
				String regex = "^(-)?\\d+$";
				if (columns.indexOf("id")>=0&&columns.indexOf("fingerprint")>0) {
					long idkey = Long.parseLong((String) map.get(columns.get(columns.indexOf("id"))));
					String fingStr = (String) map.get(columns.get(columns.indexOf("fingerprint")));
					if (fingStr != null && !fingStr.equals("")
							&& fingStr.matches(regex)) {
						long fingkey = Long.parseLong(fingStr);
						if (idmap.get(idkey) == null
								&& fingmap.get(fingkey) == null) {
							i++;
							for (int j = 0; j < columns.size(); j++) {
								String value = (String) map.get(columns.get(j));
								if (value == null || value.equals("null")
										|| value.equals("NULL")) {
									value = "";
								}
								row.createCell((short) j).setCellValue(value);
								
								if (columns.get(j).equals("titlewithurl")) {
									try {
										// 创建超链接
										XSSFCell cell = row.getCell(columns.indexOf("title"));
										String formula = "HYPERLINK(\""
												+ map.get("url")
												+ "\",\""
												+ StringUtils
														.trimToEmpty(
																""
																		+ map.get("title"))
														.replace("\"", "")
														.replace("'", "") + "\")";
										cell.setCellFormula(formula);
									} catch (Exception ex) {
										
									}
								}
								
								
							}
							idmap.put(idkey, idkey);
							fingmap.put(fingkey, fingkey);
						}
					}
				} else if(columns.indexOf("id")>=0&&columns.indexOf("titlefingerprint")>0){
					long idkey = Long.parseLong((String) map.get(columns.get(columns.indexOf("id"))));
					String tfingStr = (String) map.get(columns.get(columns.indexOf("titlefingerprint")));
					if (tfingStr != null && !tfingStr.equals("")
							&& tfingStr.matches(regex)) {
						long tfingkey = Long.parseLong(tfingStr);
						if (idmap.get(idkey) == null
								&& tfingmap.get(tfingkey) == null) {
							i++;
							for (int j = 0; j < columns.size(); j++) {
								String value = (String) map.get(columns.get(j));
								if (value == null || value.equals("null")
										|| value.equals("NULL")) {
									value = "";
								}
								row.createCell((short) j).setCellValue(value);
								
								if (columns.get(j).equals("titlewithurl")) {
									try {
										// 创建超链接
										XSSFCell cell = row.getCell(columns.indexOf("title"));
										String formula = "HYPERLINK(\""
												+ map.get("url")
												+ "\",\""
												+ StringUtils
														.trimToEmpty(
																""
																		+ map.get("title"))
														.replace("\"", "")
														.replace("'", "") + "\")";
										cell.setCellFormula(formula);
									} catch (Exception ex) {
										
									}
								}
							}
							idmap.put(idkey, idkey);
							tfingmap.put(tfingkey, tfingkey);
						}
					}
				} else if(columns.indexOf("id")>=0){
					long idkey = Long.parseLong((String) map.get(columns.get(columns.indexOf("id"))));
						if (idmap.get(idkey) == null) {
							i++;
							for (int j = 0; j < columns.size(); j++) {
								String value = (String) map.get(columns.get(j));
								if (value == null || value.equals("null")
										|| value.equals("NULL")) {
									value = "";
								}
								row.createCell((short) j).setCellValue(value);
								
								if (columns.get(j).equals("titlewithurl")) {
									try {
										// 创建超链接
										XSSFCell cell = row.getCell(columns.indexOf("title"));
										String formula = "HYPERLINK(\""
												+ map.get("url")
												+ "\",\""
												+ StringUtils
														.trimToEmpty(
																""
																		+ map.get("title"))
														.replace("\"", "")
														.replace("'", "") + "\")";
										cell.setCellFormula(formula);
									} catch (Exception ex) {
										
									}
								}
							}
							idmap.put(idkey, idkey);
						}
				}
			}
		}
	}

	
	
	public void createExcelDistinct(XSSFSheet sheet,
                                    List<Map<String, String>> dateList, String[] headers,
                                    String[] columns) {
		XSSFRow row;
		if (headers != null || headers.length > 0) {
			row = sheet.createRow((int) 0);
			for (int i = 0; i < headers.length; i++) {
				String titleName = StringUtils.trimToEmpty(headers[i]);
				if (!"".equals(titleName)) {
					Cell cell = row.createCell((short) i);
					cell.setCellValue(titleName);
					// CellStyle style = wb.createCellStyle(); //设置单元格颜色
					// style.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
					// style.setFillPattern(CellStyle.SOLID_FOREGROUND);
					// cell.setCellStyle(style);
				}
			}
		}
		int rowCount = sheet.getLastRowNum();
		if (dateList != null && dateList.size() > 0) {

			int i = 0;// 行数
			for (Map map : dateList) {
				// j为列数
				row = sheet.createRow((int) rowCount + i + 1);
				long idkey = Long.parseLong((String) map.get(columns[0]));
				String fingStr = (String) map.get(columns[1]);
				String tfingStr = (String) map.get(columns[2]);
				String regex = "^(-)?\\d+$";
				if (columns[0].equals("id") && columns[1].equals("fingerprint")
						&& columns[2].equals("titlefingerprint")) {
					if (fingStr != null && !fingStr.equals("")
							&& fingStr.matches(regex) &&tfingStr!=null&& !tfingStr.equals("")
							&& tfingStr.matches(regex)) {
						long fingkey = Long.parseLong(fingStr);
						long tfingkey = Long.parseLong(tfingStr);
						if (idmap.get(idkey) == null
								&& fingmap.get(fingkey) == null
								&& tfingmap.get(tfingkey) == null) {
							i++;
							for (int j = 0; j < columns.length; j++) {
								String value = (String) map.get(columns[j]);
								if (value == null || value.equals("null")
										|| value.equals("NULL")) {
									value = "";
								}
								row.createCell((short) j).setCellValue(value);
							}
							idmap.put(idkey, idkey);
							fingmap.put(fingkey, fingkey);
							tfingmap.put(tfingkey, tfingkey);
						}
					}
				} else if(columns[0].equals("id") && columns[1].equals("fingerprint")){
					if (fingStr != null && !fingStr.equals("")
							&& fingStr.matches(regex)) {
						long fingkey = Long.parseLong(fingStr);
						if (idmap.get(idkey) == null
								&& fingmap.get(fingkey) == null) {
							i++;
							for (int j = 0; j < columns.length; j++) {
								String value = (String) map.get(columns[j]);
								if (value == null || value.equals("null")
										|| value.equals("NULL")) {
									value = "";
								}
								row.createCell((short) j).setCellValue(value);
							}
							idmap.put(idkey, idkey);
							fingmap.put(fingkey, fingkey);
						}
					}
				} else if(columns[0].equals("id")){
						if (idmap.get(idkey) == null) {
							i++;
							for (int j = 0; j < columns.length; j++) {
								String value = (String) map.get(columns[j]);
								if (value == null || value.equals("null")
										|| value.equals("NULL")) {
									value = "";
								}
								row.createCell((short) j).setCellValue(value);
							}
							idmap.put(idkey, idkey);
						}
				}
			}
		}
	}

	/**
	 * 使用map id排重
	 * 
	 * @param sheet
	 * @param dataMap
	 * @param headers
	 * @param columns
	 */
	public void createExcel(XSSFSheet sheet,
                            Map<String, Map<String, String>> dataMap, String[] headers,
                            String[] columns) {
		XSSFRow row;
		if (headers != null || headers.length > 0) {
			row = sheet.createRow((int) 0);
			for (int i = 0; i < headers.length; i++) {
				String titleName = StringUtils.trimToEmpty(headers[i]);
				if (!"".equals(titleName)) {
					Cell cell = row.createCell((short) i);
					cell.setCellValue(titleName);
					// CellStyle style = wb.createCellStyle(); //设置单元格颜色
					// style.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
					// style.setFillPattern(CellStyle.SOLID_FOREGROUND);
					// cell.setCellStyle(style);
				}
			}
		}
		int rowCount = sheet.getLastRowNum();
		if (dataMap != null && dataMap.size() > 0) {

			int i = -1;// 行数
			for (String key : dataMap.keySet()) {
				Map<String, String> map = dataMap.get(key);
				i++;
				// j为列数
				row = sheet.createRow((int) rowCount + i + 1);
				for (int j = 0; j < columns.length; j++) {
					String value = (String) map.get(columns[j]);
					if (value == null || value.equals("null")
							|| value.equals("NULL")) {
						value = "";
					}
					row.createCell((short) j).setCellValue(value);
				}
			}
		}
	}

	public void createExcel(XSSFSheet sheet,
                            List<Map<String, String>> dateList, String[] headers,
                            String[] columns) {
		XSSFRow row;
		if (headers != null || headers.length > 0) {
			row = sheet.createRow((int) 0);
			for (int i = 0; i < headers.length; i++) {
				String titleName = StringUtils.trimToEmpty(headers[i]);
				if (!"".equals(titleName)) {
					Cell cell = row.createCell((short) i);
					cell.setCellValue(titleName);
					// CellStyle style = wb.createCellStyle(); //设置单元格颜色
					// style.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
					// style.setFillPattern(CellStyle.SOLID_FOREGROUND);
					// cell.setCellStyle(style);
				}
			}
		}
		int rowCount = sheet.getLastRowNum();
		if (dateList != null && dateList.size() > 0) {

			int i = -1;// 行数
			for (Map map : dateList) {
				i++;
				// j为列数
				row = sheet.createRow((int) rowCount + i + 1);
				for (int j = 0; j < columns.length; j++) {
					String value = (String) map.get(columns[j]);
					if (value == null || value.equals("null")
							|| value.equals("NULL")) {
						value = "";
					}
					row.createCell((short) j).setCellValue(value);
				}
			}
		}
	}

	public void createExcelChangValue(XSSFSheet sheet,
                                      List<Map<String, String>> dateList, String[] headers,
                                      String[] columns) throws IOException {
		init1();
		init2();
		XSSFRow row;
		if (headers != null || headers.length > 0) {
			row = sheet.createRow((int) 0);
			for (int i = 0; i < headers.length; i++) {
				String titleName = StringUtils.trimToEmpty(headers[i]);
				if (!"".equals(titleName)) {
					Cell cell = row.createCell((short) i);
					cell.setCellValue(titleName);
					// CellStyle style = wb.createCellStyle(); //设置单元格颜色
					// style.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
					// style.setFillPattern(CellStyle.SOLID_FOREGROUND);
					// cell.setCellStyle(style);
				}
			}
		}
		int rowCount = sheet.getLastRowNum();
		String cellvalue;
		XSSFCell cell;
		StringBuffer sb;
		if (dateList != null && dateList.size() > 0) {

			int i = -1;// 行数
			for (Map map : dateList) {
				i++;
				// j为列数
				row = sheet.createRow((int) rowCount + i + 1);
				for (int j = 0; j < columns.length; j++) {
					String value = (String) map.get(columns[j]);
					if (value == null || value.equals("null")
							|| value.equals("NULL")) {
						value = "";
					}
					row.createCell((short) j).setCellValue(value);
					if (j == 7) {
						cell = row.getCell(j);
						cellvalue = cell.toString();
						if (cellvalue != null && !cellvalue.trim().equals("")) {
							sb = new StringBuffer();
							sb.append("");
							String[] split = cellvalue.split(",");
							for (String key : split) {
								sb.append(areaMap.get(key) == null ? ""
										: areaMap.get(key) + ",");
							}
							cell.setCellValue(sb.toString());
							sb = new StringBuffer();
							sb.append("");
						}
					}
					if (j == 8) {
						cell = row.getCell(j);
						cellvalue = cell.toString();
						if (cellvalue != null && !cellvalue.trim().equals("")) {
							sb = new StringBuffer();
							sb.append("");
							String[] split = cellvalue.split(",");
							for (String key : split) {
								sb.append(sortMap.get(key) == null ? ""
										: sortMap.get(key));
							}
							cell.setCellValue(sb.toString());
							sb = new StringBuffer();
							sb.append("");
						}
					}
					if (j == 9) {
						cell = row.getCell(j);
						cellvalue = cell.toString();
						if (cellvalue != null && !cellvalue.trim().equals("")) {
							sb = new StringBuffer();
							sb.append("");
							String[] split = cellvalue.split(",");
							for (String key : split) {
								int articletype = Integer.parseInt(key);
								switch (articletype) {
								case 1:
									sb.append("新闻");
									break;
								case 2:
									sb.append("论坛");
									break;
								case 3:
									sb.append("博客");
									break;
								case 4:
									sb.append("报刊");
									break;
								case 5:
									sb.append("APP");
									break;
								case 9:
									sb.append("微信");
									break;
								default:
									sb.append("");
									break;
								}

							}
							cell.setCellValue(sb.toString());
							sb = new StringBuffer();
							sb.append("");
						}
					}
				}
			}
		}
	}

	public void outToFile() {
		try {
			wb.write(out);
			out.close();
			wb = null;
		} catch (IOException e) {
			System.out.println("输出文件异常");
			e.printStackTrace();
		}
	}

	private static Map<String, String> areaMap = new HashMap<String, String>();
	private static Map<String, String> sortMap = new HashMap<String, String>();

	public static void init1() throws IOException {
		FileReader reader = new FileReader("F://area.csv");
		BufferedReader br = new BufferedReader(reader);
		String str = br.readLine();
		while ((str = br.readLine()) != null) {
			String[] split = str.split(",");
			if (split.length != 2) {
				return;
			}
			areaMap.put(split[0], split[1]);
		}
	}

	public static void init2() throws IOException {
		FileReader reader = new FileReader("F://sort.csv");
		BufferedReader br = new BufferedReader(reader);
		String str = br.readLine();
		while ((str = br.readLine()) != null) {
			String[] split = str.split(",");
			if (split.length != 3) {
				return;
			}
			// System.out.println(split[0]+split[1]);
			sortMap.put(split[0], split[1]);
		}
	}

}
