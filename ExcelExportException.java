package com.dscanalytics.demand.model.excelexport;

/**
 * Exceção personalizada para erros de exportação.
 */
public class ExcelExportException extends RuntimeException {

	private static final long serialVersionUID = 1L;

	public ExcelExportException(String message) {
		super(message);
	}

	public ExcelExportException(String message, Throwable cause) {
		super(message, cause);
	}
}