package com.demo.exceldemomvn.service;

public interface FileConversion {

	/**
	 * 1.生成模版文件的实例
	 * 2.根据源文件同模版文件的映射关系 进行每个sheet页数据内容的写入
	 * 3.多个sheet页组合成为目标文件
	 * @param srcFile
	 * @param destinationFile
	 * @throws Exception
	 */
	void conversion(String srcFile, String destinationFile,String dataDate) throws Exception;

}
