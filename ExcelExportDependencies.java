package com.dscanalytics.demand.model.excelexport;

import java.util.Objects;

import org.springframework.data.mongodb.core.MongoTemplate;

import com.dscanalytics.demand.service.MonitoringService;
import com.dscanalytics.demand.service.OrgService;

import lombok.Getter;
import lombok.Setter;
import lombok.experimental.SuperBuilder;

//=============================================================================
//CLASSE DE DEPENDÊNCIAS
//=============================================================================

/**
 * Dependências externas necessárias para importação.
 */
@SuperBuilder
@Getter
@Setter
public class ExcelExportDependencies {
	private final MongoTemplate mongoTemplate;
	private final OrgService orgService;
	private final MonitoringService monitoringService;

	/**
	 * Construtor das dependências.
	 */
	public ExcelExportDependencies(MongoTemplate mongoTemplate, OrgService orgService, MonitoringService monitoringService) {
		this.mongoTemplate = Objects.requireNonNull(mongoTemplate);
		this.orgService = Objects.requireNonNull(orgService);
		this.monitoringService = Objects.requireNonNull(monitoringService);
	}

}
