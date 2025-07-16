package com.dscanalytics.demand.model.excelexport;

import java.util.function.Function;

import org.apache.poi.ss.usermodel.CellStyle;
import org.springframework.beans.BeanWrapperImpl;

import com.dscanalytics.demand.dto.ColumnByExportAndImportDTO;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;
import lombok.experimental.SuperBuilder;
import lombok.extern.log4j.Log4j2;

/**
 * Mapeia uma coluna do DTO para uma coluna do Excel.
 * 
 * @param <Output> Tipo dos dados
 */

@SuperBuilder
@AllArgsConstructor
@Getter
@Setter
@Log4j2
public class ExcelExportColumnMapping<Output> {

	private final int index;
	private final String header;
	private final CellStyle dataStyle;
	private final boolean rowSpan;
	private final Function<Output, Object> valueExtractor;

	/**
	 * Construtor do mapeamento de coluna.
	 */
	ExcelExportColumnMapping(int index, ColumnByExportAndImportDTO dto, Class<Output> dataClass, ExcelExportStyleManager styleManager) {
		this.index = index;
		this.header = dto.getColumnName();
		this.dataStyle = styleManager.getDataStyle(dto);
		this.rowSpan = Boolean.TRUE.equals(dto.getRowSpan());
		this.valueExtractor = createValueExtractor(dto.getField(), dataClass);
	}

	/**
	 * Cria um extrator de valor usando reflection com cache.
	 * 
	 * @param fieldPath Caminho do campo
	 * @param dataClass Classe dos dados
	 * @return Função extratora de valor
	 */
	private Function<Output, Object> createValueExtractor(String fieldPath, Class<Output> dataClass) {
		return item -> {
			try {
				BeanWrapperImpl wrapper = new BeanWrapperImpl(item);
				wrapper.setAutoGrowNestedPaths(true);
				return wrapper.getPropertyValue(fieldPath);

			} catch (Exception e) {
				log.debug("Erro ao extrair valor do campo '{}' - retornando null", fieldPath, e);
				return null;
			}
		};
	}
}
