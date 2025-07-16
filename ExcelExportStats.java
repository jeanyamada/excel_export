package com.dscanalytics.demand.model.excelexport;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

/**
 * Estatísticas da exportação.
 */

@AllArgsConstructor
@NoArgsConstructor
@Getter
@Setter
public class ExcelExportStats {
	private int totalProcessed = 0;

	public void addInProcessed(int count) {
		totalProcessed += count;
	}

}