package com.dscanalytics.demand.model.excelexport;

import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import com.dscanalytics.demand.dto.ColumnByExportAndImportDTO;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;
import lombok.experimental.SuperBuilder;
import lombok.extern.log4j.Log4j2;

/**
 * Configuração completa de uma planilha.
 * 
 * @param <Output> Tipo dos dados
 */

@SuperBuilder
@AllArgsConstructor
@Getter
@Setter
@Log4j2
public class ExcelExportSheetConfiguration<Output> {

	private final Class<Output> dataClass;
	private final String sheetName;
	private final List<ColumnByExportAndImportDTO> columnDTOs;
	private final int freezePaneColumns;
	private final int freezePaneRows;
	private final int rowSpanQuantity;

	private SXSSFSheet sheet;
	private List<ExcelExportColumnMapping<Output>> columnMappings;
	private final AtomicInteger currentRow = new AtomicInteger(1); // Começa em 1 para pular cabeçalho

	/**
	 * Construtor da configuração da planilha.
	 */
	ExcelExportSheetConfiguration(Class<Output> dataClass, String sheetName, List<ColumnByExportAndImportDTO> columnDTOs,
			int freezePaneColumns, int freezePaneRows, int rowSpanQuantity) {
		this.dataClass = dataClass;
		this.sheetName = sheetName;
		this.columnDTOs = new ArrayList<>(columnDTOs);
		this.freezePaneColumns = freezePaneColumns;
		this.freezePaneRows = freezePaneRows;
		this.rowSpanQuantity = rowSpanQuantity;
	}

	/**
	 * Inicializa a planilha com cabeçalhos e estilos.
	 * 
	 * @param workbook Workbook do Excel
	 */
	void initialize(SXSSFWorkbook workbook) {
		try {
			this.sheet = workbook.createSheet(sheetName);
			ExcelExportStyleManager styleManager = new ExcelExportStyleManager(workbook);

			this.columnMappings = new ArrayList<>();
			Row headerRow = sheet.createRow(0);

			// Cria cabeçalhos e mapeamentos de colunas
			for (int i = 0; i < columnDTOs.size(); i++) {
				ColumnByExportAndImportDTO dto = columnDTOs.get(i);

				// Cria mapeamento da coluna
				columnMappings.add(new ExcelExportColumnMapping<>(i, dto, dataClass, styleManager));

				// Cria célula do cabeçalho
				Cell headerCell = headerRow.createCell(i);
				headerCell.setCellValue(dto.getColumnName());
				headerCell.setCellStyle(styleManager.getHeaderStyle(dto));
			}

			// Aplica congelamento de painéis
			sheet.createFreezePane(freezePaneColumns, freezePaneRows);

			log.debug("Planilha '{}' inicializada com {} colunas", sheetName, columnDTOs.size());

		} catch (Exception e) {
			log.error("Erro ao inicializar planilha", e);
			// throw new ExcelExportException("Erro ao inicializar planilha", e);
		}
	}

	public int getCurrentRowIndex() {
		return currentRow.getAndIncrement();
	}
}