package com.dscanalytics.demand.model.excelexport;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.function.Function;

import com.dscanalytics.demand.util.Util;

import lombok.Getter;
import lombok.Setter;
import lombok.experimental.SuperBuilder;
import lombok.extern.log4j.Log4j2;

/**
 * Registro de mapeadores de tipos de valores.
 */

@SuperBuilder
@Getter
@Setter
@Log4j2
public class ExcelExportValueMapperRegistry {

	private final Map<Class<?>, Function<Object, ?>> mappers = new HashMap<>();

	/**
	 * Construtor que registra os mapeadores padrão.
	 */
	ExcelExportValueMapperRegistry() {
		registerDefaultMappers();
	}

	/**
	 * Registra os mapeadores padrão para tipos comuns.
	 */
	private void registerDefaultMappers() {
		mappers.put(Date.class, val -> val);
		mappers.put(LocalDate.class, val -> convertLocalDateToDate((LocalDate) val));
		mappers.put(LocalDateTime.class, val -> convertLocalDateTimeToDate((LocalDateTime) val));
	}

	/**
	 * Converte LocalDate para Date.
	 */
	private Date convertLocalDateToDate(LocalDate localDate) {
		try {
			return Util.asDate(localDate);
		} catch (Exception e) {
			log.warn("Erro ao converter LocalDate para Date: {}", localDate, e);
			return null;
		}
	}

	/**
	 * Converte LocalDateTime para Date.
	 */
	private Date convertLocalDateTimeToDate(LocalDateTime localDateTime) {
		try {
			return Util.asDate(localDateTime);
		} catch (Exception e) {
			log.warn("Erro ao converter LocalDateTime para Date: {}", localDateTime, e);
			return null;
		}
	}

	/**
	 * Obtém os mapeadores registrados.
	 * 
	 * @return Mapa de mapeadores
	 */
	Map<Class<?>, Function<Object, ?>> getMappers() {
		return new HashMap<>(mappers);
	}
}
