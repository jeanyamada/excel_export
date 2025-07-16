package com.dscanalytics.demand.model.excelexport;

import java.util.stream.Stream;

import com.dscanalytics.demand.dto.ExportAndImportConfigDTO;
import com.dscanalytics.demand.dto.OrgDTO;
import com.dscanalytics.demand.dto.ScenarioDTO;
import com.dscanalytics.demand.dto.UserDTO;

import lombok.Getter;
import lombok.Setter;
import lombok.experimental.SuperBuilder;

//=============================================================================
//CLASSES DE DADOS
//=============================================================================

/**
 * Requisição de exportação com todos os dados necessários.
 */
@SuperBuilder
@Getter
@Setter
public class ExcelExportRequest<Output> {
	private final Stream<Output> dataStream;
	private final Long totalRecords;
	private final String correlationId;
	private final String monitoringId;
	private final Object find;
	private final ExportAndImportConfigDTO exportAndImportConfig;
	private final OrgDTO org;
	private final UserDTO user;
	private final ScenarioDTO scenario;

	/**
	 * Construtor da requisição.
	 */
	public ExcelExportRequest(Stream<Output> dataStream, Long totalRecords, String correlationId, String monitoringId,
			Object find, ExportAndImportConfigDTO exportAndImportConfig, OrgDTO org, UserDTO user,
			ScenarioDTO scenario) {
		super();
		this.dataStream = dataStream;
		this.totalRecords = totalRecords;
		this.correlationId = correlationId;
		this.monitoringId = monitoringId;
		this.find = find;
		this.exportAndImportConfig = exportAndImportConfig;
		this.org = org;
		this.user = user;
		this.scenario = scenario;
	}

}
