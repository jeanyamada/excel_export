package com.dscanalytics.demand.model.excelexport;

import java.io.ByteArrayInputStream;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Objects;
import java.util.Set;

import com.dscanalytics.demand.dto.ExportAndImportConfigDTO;
import com.dscanalytics.demand.dto.OrgDTO;
import com.dscanalytics.demand.dto.UserDTO;
import com.dscanalytics.demand.service.MonitoringService;
import com.dscanalytics.demand.util.Util;

import lombok.Getter;
import lombok.Setter;
import lombok.experimental.SuperBuilder;
import lombok.extern.log4j.Log4j2;

/**
 * Exportador de dados para Excel com processamento em lotes e otimizações de
 * performance.
 * 
 * Funcionalidades principais: - Exportação de dados em lotes para evitar
 * problemas de memória - Monitoramento de progresso em tempo real - Validação
 * de dados personalizada - Callbacks para processamento antes/depois da
 * exportação - Tratamento robusto de erros com logging detalhado - Configuração
 * flexível de parâmetros de exportação
 * 
 * @param <Output> Tipo de dados de saída para o Excel
 * 
 * @author Jean Yamada
 * @since 1.0
 */

@SuperBuilder
@Getter
@Setter
@Log4j2
public class ExcelExportByGrid<Output> {

	private static final int DEFAULT_PAGE_SIZE = 1000;
	private static final String ERROR_MESSAGE_PREFIX = "Erro durante exportação Excel";

	// Configurações básicas
	private final ExcelExportConfiguration<Output> configuration;
	private final ExcelExportDependencies dependencies;

	// Controle de erros
	private final Set<String> errors = new HashSet<>();

	/**
	 * Construtor principal do exportador.
	 * 
	 * @param configuration Configuração de exportação
	 * @param dependencies  Dependências externas
	 */
	public ExcelExportByGrid(ExcelExportConfiguration<Output> configuration, ExcelExportDependencies dependencies) {
		this.configuration = Objects.requireNonNull(configuration, "Configuração não pode ser nula");
		this.dependencies = Objects.requireNonNull(dependencies, "Dependências não podem ser nulas");

	}

	/**
	 * Executa a exportação do Excel processando os dados em lotes.
	 * 
	 * @param request Dados da requisição de exportação
	 * @return ByteArrayInputStream com o arquivo Excel gerado, ou null em caso de
	 *         erro
	 */

	public ByteArrayInputStream excelExport(ExcelExportRequest<Output> request) {
		log.info("Iniciando exportação Excel para coleção: {}", configuration.getCollection());

		try {
			// Validações iniciais
			validateRequest(request);

			// Cria o gerador do Excel
			ExcelExport<Output> excelGenerator = createExcelGenerator(request);

			// Processa os dados em lotes
			ExcelExportStats stats = processDataInBatches(request, excelGenerator);

			// Finaliza exportação
			ByteArrayInputStream result = excelGenerator.save();

			// Log de conclusão
			log.info("Exportação concluída com sucesso. Registros processados: {}, Erros: {}",
					stats.getTotalProcessed(), errors.size());

			return result;

		} catch (Exception e) {
			log.error("{} - Coleção: {}, CorrelationId: {}", ERROR_MESSAGE_PREFIX, configuration.getCollection(),
					request.getCorrelationId(), e);
			errors.add("Erro durante exportação: " + e.getMessage());
			return null;
		}
	}

	/**
	 * Valida os dados da requisição de exportação.
	 */
	private void validateRequest(ExcelExportRequest<Output> request) {
		if (request == null) {
			throw new IllegalArgumentException("Requisição não pode ser nula");
		}

		if (request.getDataStream() == null) {
			throw new IllegalArgumentException("Stream de dados não pode ser nulo");
		}

		if (request.getExportAndImportConfig() == null) {
			throw new IllegalArgumentException("Configuração de exportação/importação não pode ser nula");
		}

		if (request.getTotalRecords() == null || request.getTotalRecords() < 0) {
			throw new IllegalArgumentException("Total de registros deve ser maior ou igual a zero");
		}

	}

	/**
	 * Cria o gerador do Excel configurado.
	 */
	private ExcelExport<Output> createExcelGenerator(ExcelExportRequest<Output> request) {
		ExportAndImportConfigDTO config = request.getExportAndImportConfig();

		return ExcelExport.builder(configuration.getOutputClass()).withColumns(config.getColumns())
				.withFreezePane(config.getCountFreezeColumn(), config.getCountFreezeRow())
				.withRowSpan(config.getRowSpanQuantity()).withSheetName(config.getSheetName()).build();
	}

	/**
	 * Processa os dados em lotes para melhor performance e gerenciamento de
	 * memória.
	 */
	private ExcelExportStats processDataInBatches(ExcelExportRequest<Output> request,
			ExcelExport<Output> excelGenerator) {
		ExcelExportStats stats = new ExcelExportStats();

		// Processa em lotes usando stream
		List<Output> currentBatch = new ArrayList<>();
		Iterator<Output> iterator = request.getDataStream().iterator();

		int processedBatches = 0;
		int totalBatches = calculateTotalBatches(request.getTotalRecords());

		try {
			while (iterator.hasNext()) {
				currentBatch.add(iterator.next());

				// Processa lote quando atinge o tamanho máximo
				if (currentBatch.size() >= DEFAULT_PAGE_SIZE) {
					processBatch(currentBatch, request, excelGenerator, stats);
					currentBatch.clear();

					processedBatches++;
					updateProgress(request, processedBatches, totalBatches);
				}
			}

			// Processa último lote se não estiver vazio
			if (!currentBatch.isEmpty()) {
				processBatch(currentBatch, request, excelGenerator, stats);
				processedBatches++;
				updateProgress(request, processedBatches, totalBatches);
			}

			// Fecha o stream
			request.getDataStream().close();

		} catch (Exception e) {
			log.error("Erro durante processamento dos lotes", e);
			errors.add("Erro no processamento: " + e.getMessage());
			// throw new ExportExcelException("Falha no processamento dos dados", e);
		}

		return stats;
	}

	/**
	 * Atualiza o progresso do monitoramento da exportação de dados do Excel.
	 *
	 * <p>
	 * Esse método é utilizado durante a exportação de dados para informar o
	 * progresso ao serviço de monitoramento, com base no número total de lotes e na
	 * quantidade processada até o momento.
	 *
	 * @param request      objeto da requisição de exportação, contendo dados da
	 *                     organização, usuário e identificadores.
	 * @param currentBatch lote atual sendo processado.
	 * @param totalBatches número total de lotes a serem processados.
	 */
	private void updateProgress(ExcelExportRequest<Output> request, int currentBatch, int totalBatches) {
		try {
			// Dados da organização e usuário
			OrgDTO org = request.getOrg();
			UserDTO user = request.getUser();

			// Dependências para monitoramento
			MonitoringService monitoringService = dependencies.getMonitoringService();
			String correlationId = request.getCorrelationId();
			String monitoringId = request.getMonitoringId();

			// Só atualiza progresso se houver dados necessários
			if (totalBatches > 0 && !Util.isNull(monitoringService) && !Util.isNullOrEmpty(monitoringId)
					&& !Util.isNullOrEmpty(correlationId)) {

				// Atualiza o progresso no serviço de monitoramento
				monitoringService.updateProgress(org, user, correlationId, monitoringId, currentBatch, totalBatches);

				if (configuration.isDebugEnabled()) {
					log.debug("Progresso atualizado: {}/{} para correlationId: {}", currentBatch, totalBatches,
							correlationId);
				}
			}

		} catch (Exception e) {
			log.error("Erro ao atualizar o progresso da exportação: {}", e.getMessage(), e);
		}
	}

	/**
	 * Processa um lote de dados aplicando validações e transformações.
	 */
	private void processBatch(List<Output> batch, ExcelExportRequest<Output> request,
			ExcelExport<Output> excelGenerator, ExcelExportStats stats) {
		if (batch.isEmpty()) {
			return;
		}

		log.debug("Processando lote de {} itens", batch.size());

		try {

			// Adiciona ao Excel
			excelGenerator.addData(batch);
			stats.addInProcessed(batch.size());

		} catch (Exception e) {
			log.error("Erro ao processar lote", e);
			errors.add("Erro ao processar lote: " + e.getMessage());
		}
	}

	/**
	 * Calcula o número total de lotes baseado no total de registros.
	 */
	private int calculateTotalBatches(Long totalRecords) {
		return (int) Math.ceil((double) totalRecords / DEFAULT_PAGE_SIZE);
	}

}
