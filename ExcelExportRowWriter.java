package com.dscanalytics.demand.model.excelexport;

import java.lang.reflect.Field;
import java.util.Objects;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import lombok.Getter;
import lombok.Setter;
import lombok.experimental.SuperBuilder;
import lombok.extern.log4j.Log4j2;

/**
 * <p>
 * Responsável por escrever uma linha de dados em uma planilha do Excel.
 * </p>
 * <p>
 * Esta classe encapsula a lógica de criação de uma linha, processamento de cada
 * coluna com base nos mapeamentos fornecidos e verificação de condições
 * especiais, como ocultar a linha.
 * </p>
 *
 * @param <Output> O tipo do objeto de dados que será escrito na linha.
 */
@SuperBuilder
@Getter
@Setter
@Log4j2
class ExcelExportRowWriter<Output> {

	private static final String HIDE_FIELD_NAME = "hide";

	private final ExcelExportSheetConfiguration<Output> sheetConfig;
	private final ExcelExportCellWriter cellWriter;

	/**
	 * Construtor para {@code ExcelExportRowWriter}.
	 *
	 * @param sheetConfig A {@link ExcelExportSheetConfiguration} que contém as configurações
	 *                    da planilha e os mapeamentos de coluna. Não pode ser nula.
	 * @param cellWriter  O {@link ExcelExportCellWriter} para definir os valores das células.
	 *                    Não pode ser nulo.
	 */
	public ExcelExportRowWriter(ExcelExportSheetConfiguration<Output> sheetConfig, ExcelExportCellWriter cellWriter) {
		this.sheetConfig = Objects.requireNonNull(sheetConfig, "ExcelExportSheetConfiguration não pode ser nula.");
		this.cellWriter = Objects.requireNonNull(cellWriter, "ExcelExportCellWriter não pode ser nulo.");
	}

	/**
	 * Adiciona um item de dados como uma nova linha na planilha do Excel.
	 *
	 * @param item O objeto de dados a ser escrito na linha. Não pode ser nulo.
	 * @throws ExcelExportException se ocorrer um erro ao adicionar a linha.
	 */
	public void addRow(Output item) throws ExcelExportException {
		Objects.requireNonNull(item, "O item de dados não pode ser nulo.");

		try {
			Row row = sheetConfig.getSheet().createRow(sheetConfig.getCurrentRowIndex());

			for (ExcelExportColumnMapping<Output> mapping : sheetConfig.getColumnMappings()) {
				processColumn(item, row, mapping);
			}

			checkAndHideRow(item, row);

		} catch (Exception e) {
			log.error("Erro ao adicionar linha para o item: {}. Erro: {}", item, e.getMessage(), e);
			throw new ExcelExportException("Falha ao adicionar linha à planilha.", e);
		}
	}

	/**
	 * Processa uma coluna específica de uma linha, extraindo o valor do item de
	 * dados e definindo-o na célula correspondente.
	 *
	 * @param item    O objeto de dados.
	 * @param row     A {@link Row} do Excel.
	 * @param mapping O {@link ExcelExportColumnMapping} para a coluna atual.
	 */
	private void processColumn(Output item, Row row, ExcelExportColumnMapping<Output> mapping) {
		Cell cell = row.createCell(mapping.getIndex());

		try {
			Object value = mapping.getValueExtractor().apply(item);
			cellWriter.setCellValue(cell, value);
			cell.setCellStyle(mapping.getDataStyle());

		} catch (Exception e) {
			log.warn("Erro ao extrair valor para a coluna \"{}\". A célula será deixada em branco. Erro: {}",
					mapping.getHeader(), e.getMessage());
			cell.setBlank();
			cell.setCellStyle(mapping.getDataStyle());
		}
	}

	/**
	 * Verifica se o item de dados possui um campo chamado "hide" com valor
	 * {@code true}. Se for o caso, a linha correspondente no Excel será ocultada.
	 *
	 * @param item O objeto de dados.
	 * @param row  A {@link Row} do Excel.
	 */
	private void checkAndHideRow(Output item, Row row) {
		try {
			Field hideField = item.getClass().getDeclaredField(HIDE_FIELD_NAME);
			hideField.setAccessible(true);

			Object hideValue = hideField.get(item);
			if (Boolean.TRUE.equals(hideValue)) {
				row.setZeroHeight(true);
				log.debug("Linha {} ocultada com base no campo 'hide'.", row.getRowNum());
			}

		} catch (NoSuchFieldException e) {
			// O campo 'hide' não existe, o que é normal. Nenhuma ação é necessária.
		} catch (Exception e) {
			log.warn("Erro ao verificar o campo 'hide' para a linha {}: {}", row.getRowNum(), e.getMessage(), e);
		}
	}
}
