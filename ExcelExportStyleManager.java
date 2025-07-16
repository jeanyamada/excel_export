package com.dscanalytics.demand.model.excelexport;

import java.util.HashMap;
import java.util.Map;
import java.util.Optional;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.util.StringUtils;

import com.dscanalytics.demand.dto.ColumnByExportAndImportDTO;

import lombok.Getter;
import lombok.Setter;
import lombok.experimental.SuperBuilder;
import lombok.extern.log4j.Log4j2;

/**
 * Gerencia a criação e cache de estilos de células.
 */
@SuperBuilder
@Getter
@Setter
@Log4j2
public class ExcelExportStyleManager {

	private final Workbook workbook;
	private final Map<String, CellStyle> styleCache = new HashMap<>();

	/**
	 * Construtor do gerenciador de estilos.
	 * 
	 * @param workbook Workbook do Excel
	 */
	ExcelExportStyleManager(Workbook workbook) {
		this.workbook = workbook;
	}

	/**
	 * Obtém o estilo do cabeçalho para uma coluna.
	 * 
	 * @param dto Configuração da coluna
	 * @return Estilo do cabeçalho
	 */
	CellStyle getHeaderStyle(ColumnByExportAndImportDTO dto) {
		String cacheKey = "header_" + dto.getColumnName();

		return styleCache.computeIfAbsent(cacheKey, k -> {
			try {
				CellStyle style = workbook.createCellStyle();
				Font font = workbook.createFont();

				font.setBold(true);
				font.setColor(Optional.ofNullable(dto.getHeaderFontColor()).orElse(IndexedColors.BLUE).getIndex());

				style.setFont(font);

				if (dto.getHeaderBackgroundColor() != null) {
					style.setFillForegroundColor(dto.getHeaderBackgroundColor().getIndex());
					style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				}

				return style;

			} catch (Exception e) {
				log.warn("Erro ao criar estilo do cabeçalho para coluna: {}", dto.getColumnName(), e);
				return workbook.createCellStyle(); // Retorna estilo padrão
			}
		});
	}

	/**
	 * Obtém o estilo dos dados para uma coluna.
	 * 
	 * @param dto Configuração da coluna
	 * @return Estilo dos dados
	 */
	public CellStyle getDataStyle(ColumnByExportAndImportDTO dto) {
		String cacheKey = "data_" + dto.getColumnName();

		return styleCache.computeIfAbsent(cacheKey, k -> {
			try {
				CellStyle style = workbook.createCellStyle();
				Font font = workbook.createFont();

				if (dto.getFontColor() != null) {
					font.setColor(dto.getFontColor().getIndex());
				}

				if (Boolean.TRUE.equals(dto.getFontBold())) {
					font.setBold(true);
				}

				style.setFont(font);

				if (StringUtils.hasText(dto.getFormat())) {
					style.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat(dto.getFormat()));
				}

				if (dto.getBackgroundColor() != null) {
					style.setFillForegroundColor(dto.getBackgroundColor().getIndex());
					style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				}

				if (Boolean.TRUE.equals(dto.getRowSpan())) {
					style.setVerticalAlignment(VerticalAlignment.CENTER);
				}

				return style;

			} catch (Exception e) {
				log.warn("Erro ao criar estilo dos dados para coluna: {}", dto.getColumnName(), e);
				return workbook.createCellStyle(); // Retorna estilo padrão
			}
		});
	}
}