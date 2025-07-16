package com.dscanalytics.demand.model.excelexport;

import java.util.ArrayList;
import java.util.List;

import org.springframework.util.StringUtils;

import com.dscanalytics.demand.dto.ColumnByExportAndImportDTO;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;

// =================================================================================
// Classes CsvExportBuilder e de Configuração
// =================================================================================

/**
 * CsvExportBuilder para configurar e criar uma instância de ExcelExport.
 * 
 * @param <Output> Tipo dos dados
 */

@AllArgsConstructor
@Getter
@Setter
public class ExcelExportBuilder<Output> {
	private final Class<Output> dataClass;
	private final List<ColumnByExportAndImportDTO> columnConfigs = new ArrayList<>();
	private String sheetName = "Dados";
	private int freezePaneColumns = 0;
	private int freezePaneRows = 1;
	private int rowSpanQuantity = 1;

	/**
	 * Construtor do CsvExportBuilder.
	 * 
	 * @param dataClass Classe dos dados
	 */
	public ExcelExportBuilder(Class<Output> dataClass) {
		this.dataClass = dataClass;
	}

	/**
	 * Define o nome da planilha.
	 * 
	 * @param sheetName Nome da planilha
	 * @return Este builder
	 */
	public ExcelExportBuilder<Output> withSheetName(String sheetName) {
		this.sheetName = StringUtils.hasText(sheetName) ? sheetName : "Dados";
		return this;
	}

	/**
	 * Define as configurações das colunas.
	 * 
	 * @param configs Lista de configurações de colunas
	 * @return Este builder
	 * @throws IllegalArgumentException se configs for null ou vazio
	 */
	public ExcelExportBuilder<Output> withColumns(List<ColumnByExportAndImportDTO> configs) {
		if (configs == null || configs.isEmpty()) {
			throw new IllegalArgumentException("Lista de configurações de colunas não pode ser nula ou vazia");
		}

		this.columnConfigs.clear();
		this.columnConfigs.addAll(configs);
		return this;
	}

	/**
	 * Define o congelamento de painéis.
	 * 
	 * @param columns Número de colunas a congelar
	 * @param rows    Número de linhas a congelar
	 * @return Este builder
	 */
	public ExcelExportBuilder<Output> withFreezePane(int columns, int rows) {
		this.freezePaneColumns = Math.max(0, columns);
		this.freezePaneRows = Math.max(0, rows);
		return this;
	}

	/**
	 * Define a quantidade de linhas para rowspan.
	 * 
	 * @param quantity Quantidade de linhas
	 * @return Este builder
	 */
	public ExcelExportBuilder<Output> withRowSpan(int quantity) {
		this.rowSpanQuantity = Math.max(1, quantity);
		return this;
	}

	/**
	 * Constrói a instância do ExcelExport.
	 * 
	 * @return Nova instância configurada
	 * @throws IllegalStateException se nenhuma coluna foi configurada
	 */
	public ExcelExport<Output> build() {
		if (columnConfigs.isEmpty()) {
			throw new IllegalStateException("Pelo menos uma coluna deve ser configurada");
		}

		ExcelExportSheetConfiguration<Output> config = new ExcelExportSheetConfiguration<>(dataClass, sheetName, columnConfigs,
				freezePaneColumns, freezePaneRows, rowSpanQuantity);

		ExcelExport<Output> exporter = new ExcelExport<>(config);

		config.initialize(exporter.getWorkbook());

		return exporter;
	}
}
