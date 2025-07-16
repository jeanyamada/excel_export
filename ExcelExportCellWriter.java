package com.dscanalytics.demand.model.excelexport;

import java.util.Calendar;
import java.util.Date;
import java.util.Map;
import java.util.Objects;
import java.util.function.Function;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.RichTextString;

import lombok.Getter;
import lombok.Setter;
import lombok.experimental.SuperBuilder;
import lombok.extern.log4j.Log4j2;

/**
 * <p>
 * Responsável por definir o valor de uma célula do Excel com base no tipo do
 * objeto fornecido.
 * </p>
 * <p>
 * Esta classe encapsula a lógica de conversão de tipos de dados Java para os
 * tipos de células do Apache POI, utilizando mapeadores de valor customizados
 * quando disponíveis para tipos específicos.
 * </p>
 */

@SuperBuilder
@Getter
@Setter
@Log4j2
class ExcelExportCellWriter {

	private final Map<Class<?>, Function<Object, ?>> valueMappers;

	/**
	 * Construtor para {@code ExcelExportCellWriter}.
	 *
	 * @param valueMappers Um {@link Map} de {@link Function}s que mapeiam objetos
	 *                     de um tipo específico para um valor que pode ser
	 *                     diretamente escrito em uma célula do Excel. Não pode ser
	 *                     nulo.
	 */
	public ExcelExportCellWriter(Map<Class<?>, Function<Object, ?>> valueMappers) {
		this.valueMappers = Objects.requireNonNull(valueMappers, "ValueMappers não pode ser nulo.");
	}

	/**
	 * Define o valor de uma {@link Cell} do Excel com base no {@code value}
	 * fornecido. O método tenta inferir o tipo do valor e aplicar o mapeamento e a
	 * escrita apropriados.
	 *
	 * @param cell  A {@link Cell} do Apache POI onde o valor será definido. Não
	 *              pode ser nula.
	 * @param value O objeto cujo valor será escrito na célula. Pode ser nulo.
	 */
	public void setCellValue(Cell cell, Object value) {
		Objects.requireNonNull(cell, "Célula não pode ser nula.");

		if (value == null) {
			cell.setBlank();
			return;
		}

		// Aplica mapeamento de tipo se disponível
		Function<Object, ?> mapper = valueMappers.get(value.getClass());
		if (mapper != null) {
			value = mapper.apply(value);
		}

		// Define o valor baseado no tipo
		try {
			if (value instanceof String) {
				cell.setCellValue((String) value);
			} else if (value instanceof Number) {
				cell.setCellValue(((Number) value).doubleValue());
			} else if (value instanceof Boolean) {
				cell.setCellValue((Boolean) value);
			} else if (value instanceof Date) {
				cell.setCellValue((Date) value);
			} else if (value instanceof Calendar) {
				cell.setCellValue((Calendar) value);
			} else if (value instanceof RichTextString) {
				cell.setCellValue((RichTextString) value);
			} else {
				// Para outros tipos, converte para String por padrão
				cell.setCellValue(value.toString());
			}
		} catch (Exception e) {
			log.warn("Erro ao definir valor da célula para o tipo {}. Tentando converter para String. Valor: {}",
					value.getClass().getSimpleName(), value, e);
			// Em caso de erro, tenta definir como String para evitar falha total
			cell.setCellValue(value.toString());
		}
	}
}
