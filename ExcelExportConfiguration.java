package com.dscanalytics.demand.model.excelexport;

import java.util.Objects;

import com.dscanalytics.demand.dto.OrgDTO;
import com.dscanalytics.demand.dto.ScenarioDTO;
import com.dscanalytics.demand.dto.UserDTO;
import com.dscanalytics.demand.util.Util;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;
import lombok.experimental.SuperBuilder;

//=============================================================================
//CLASSE DE CONFIGURAÇÃO
//=============================================================================

/**
 * Configuração para importação de dados do Excel.
 */
@SuperBuilder
@AllArgsConstructor
@Getter
@Setter
public class ExcelExportConfiguration<Output> {
	private final String collection;
	private final String key;
	private final Class<Output> outputClass;
	private final OrgDTO org;
	private final UserDTO user;
	private final ScenarioDTO scenario;
	private final boolean debugEnabled;

	/**
	 * Construtor da configuração.
	 */
	public ExcelExportConfiguration(String collection, String key, Class<Output> outputClass, OrgDTO org, UserDTO user,
			ScenarioDTO s) {
		this.collection = Objects.requireNonNull(collection);
		this.key = Objects.requireNonNull(key);
		this.outputClass = Objects.requireNonNull(outputClass);
		this.org = Objects.requireNonNull(org);
		this.user = Objects.requireNonNull(user);
		this.scenario = Objects.requireNonNull(s);
		this.debugEnabled = !Util.isNull(org.getDebug()) && org.getDebug();
	}

	public ExcelExportConfiguration(String collection, String key, Class<Output> outputClass, OrgDTO org,
			UserDTO user) {
		this.collection = Objects.requireNonNull(collection);
		this.key = Objects.requireNonNull(key);
		this.outputClass = Objects.requireNonNull(outputClass);
		this.org = Objects.requireNonNull(org);
		this.user = Objects.requireNonNull(user);
		this.scenario = null;
		this.debugEnabled = !Util.isNull(org.getDebug()) && org.getDebug();
	}

}
