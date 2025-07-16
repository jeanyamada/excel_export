package com.dscanalytics.demand.model.excelexport;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.Closeable;
import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.function.Function;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import lombok.Getter;
import lombok.Setter;
import lombok.extern.log4j.Log4j2;

/**
 * Exportador de Excel genérico e altamente configurável que utiliza o padrão
 * CsvExportBuilder para facilitar a criação de planilhas complexas.
 * 
 * Esta classe oferece suporte para: - Processamento em streaming para grandes
 * volumes de dados - Estilos customizáveis para cabeçalhos e dados - Mesclagem
 * de células (rowspan) - Formatação automática de tipos de dados -
 * Gerenciamento eficiente de memória - Tratamento robusto de erros
 * 
 * 
 * @param <Output> O tipo do objeto de dados a ser exportado
 * 
 * @author Jean Yamada
 * @version 2.0
 * @since 1.0
 */

@Getter
@Setter
@Log4j2
public class ExcelExport<Output> implements Closeable {

	/**
	 * Configuração padrão para o número de linhas mantidas na memória antes de
	 * serem escritas no disco. -1 desativa o auto-flush e usa o disco para janelas
	 * de linhas, ideal para grandes arquivos.
	 */
	private static final int STREAMING_MEMORY_THRESHOLD = -1;

	private final SXSSFWorkbook workbook;
	private final ExcelExportSheetConfiguration<Output> sheetConfig;
	private final Map<Class<?>, Function<Object, ?>> valueMappers;
	private final ExcelExportValueMapperRegistry valueMapperRegistry;
	private final ExcelExportRowWriter<Output> rowWriter;
	private final ExcelExportCellWriter cellWriter;
	private volatile boolean isClosed = false;

	/**
	 * Construtor privado - use o CsvExportBuilder para criar instâncias.
	 * 
	 * @param sheetConfig Configuração da planilha
	 */
	ExcelExport(ExcelExportSheetConfiguration<Output> sheetConfig) {
		this.workbook = new SXSSFWorkbook(STREAMING_MEMORY_THRESHOLD);
		this.sheetConfig = Objects.requireNonNull(sheetConfig, "Configuração da planilha não pode ser nula");
		this.valueMapperRegistry = new ExcelExportValueMapperRegistry();
		this.valueMappers = valueMapperRegistry.getMappers();
		this.cellWriter = new ExcelExportCellWriter(valueMappers);
		this.rowWriter = new ExcelExportRowWriter<>(sheetConfig, cellWriter);
		log.debug("ExcelExport inicializado para classe: {}, planilha: {}", sheetConfig.getDataClass().getSimpleName(),
				sheetConfig.getSheetName());
	}

	/**
	 * Ponto de entrada para criar uma nova exportação.
	 * 
	 * @param dataClass A classe dos objetos de dados (ex: MyDataObject.class)
	 * @param <Output>  Tipo dos dados
	 * @return Um builder para configurar a planilha
	 * @throws IllegalArgumentException se dataClass for null
	 */
	public static <Output> ExcelExportBuilder<Output> builder(Class<Output> dataClass) {
		if (dataClass == null) {
			throw new IllegalArgumentException("Classe de dados não pode ser nula");
		}
		return new ExcelExportBuilder<>(dataClass);
	}

	/**
	 * Adiciona uma lista de dados à planilha.
	 * 
	 * @param data A lista de objetos a serem escritos
	 * @throws IllegalStateException    se o exportador estiver fechado
	 * @throws IllegalArgumentException se data for null
	 * @throws ExcelExportException     se ocorrer erro durante a adição dos dados
	 */
	public void addData(List<Output> data) {
		validateNotClosed();

		if (data == null) {
			throw new IllegalArgumentException("Lista de dados não pode ser nula");
		}

		if (data.isEmpty()) {
			log.debug("Lista de dados vazia - nenhum dado adicionado");
			return;
		}

		log.debug("Adicionando {} registros à planilha", data.size());

		try {
			for (Output item : data) {
				rowWriter.addRow(item);
			}
			log.debug("Todos os {} registros foram adicionados com sucesso", data.size());

		} catch (Exception e) {
			log.error("Erro ao adicionar dados à planilha", e);
			// throw new ExcelExportException("Falha ao adicionar dados à planilha", e);
		}
	}

	/**
	 * Finaliza o processo de exportação, aplica formatações finais como mesclagem e
	 * auto-ajuste de colunas, e retorna o arquivo como um ByteArrayInputStream.
	 * 
	 * @return Um InputStream com o conteúdo do arquivo Excel
	 * @throws IllegalStateException se o exportador estiver fechado
	 * @throws ExcelExportException  se ocorrer erro durante a finalização
	 */
	public ByteArrayInputStream save() {
		validateNotClosed();

		log.debug("Iniciando processo de finalização da exportação");

		try {
			applyFinalFormatting();

			ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
			workbook.write(outputStream);

			log.debug("Exportação finalizada com sucesso");
			return new ByteArrayInputStream(outputStream.toByteArray());

		} catch (IOException e) {
			log.error("Erro de IO ao salvar arquivo Excel", e);
			// throw new ExcelExportException("Erro ao gerar arquivo Excel", e);

		} catch (Exception e) {
			log.error("Erro inesperado ao salvar arquivo Excel", e);
			// throw new ExcelExportException("Erro inesperado durante a exportação", e);

		} finally {
			try {
				close();

			} catch (IOException e) {
				log.warn("Erro ao fechar recursos após salvamento", e);
			}
		}

		return null;
	}

	/**
	 * Aplica formatações finais à planilha.
	 */
	private void applyFinalFormatting() {
		log.debug("Aplicando formatações finais");

		try {
			applyAutoSizeColumns();
			applyRowSpanMerging();

		} catch (Exception e) {
			log.error("Erro ao aplicar formatações finais", e);
			// throw new ExcelExportException("Erro ao aplicar formatações finais", e);
		}
	}

	/**
	 * Aplica auto-ajuste de colunas.
	 */
	private void applyAutoSizeColumns() {
		log.debug("Aplicando auto-ajuste de colunas");

		try {
			for (int i = 0; i < sheetConfig.getColumnMappings().size(); i++) {
				sheetConfig.getSheet().trackColumnForAutoSizing(i);
				sheetConfig.getSheet().autoSizeColumn(i);
			}

		} catch (Exception e) {
			log.warn("Erro ao aplicar auto-ajuste de colunas", e);
			// Não interrompe o processo se o auto-ajuste falhar
		}
	}

	/**
	 * Aplica mesclagem de células para colunas com rowspan.
	 */
	private void applyRowSpanMerging() {
		if (sheetConfig.getRowSpanQuantity() <= 1) {
			return;
		}

		log.debug("Aplicando mesclagem de células - rowspan: {}", sheetConfig.getRowSpanQuantity());

		try {
			int lastRow = sheetConfig.getSheet().getLastRowNum();

			for (ExcelExportColumnMapping<Output> mapping : sheetConfig.getColumnMappings()) {
				if (mapping.isRowSpan()) {
					mergeCellsForColumn(mapping, lastRow);
				}
			}

		} catch (Exception e) {
			log.warn("Erro ao aplicar mesclagem de células", e);
			// Não interrompe o processo se a mesclagem falhar
		}
	}

	/**
	 * Mescla células para uma coluna específica.
	 * 
	 * @param mapping Mapeamento da coluna
	 * @param lastRow Última linha da planilha
	 */
	private void mergeCellsForColumn(ExcelExportColumnMapping<Output> mapping, int lastRow) {
		for (int rowIndex = 1; rowIndex <= lastRow; rowIndex += sheetConfig.getRowSpanQuantity()) {
			int endRow = Math.min(rowIndex + sheetConfig.getRowSpanQuantity() - 1, lastRow);

			if (rowIndex < endRow) {
				try {
					CellRangeAddress mergeRange = new CellRangeAddress(rowIndex, endRow, mapping.getIndex(),
							mapping.getIndex());
					sheetConfig.getSheet().addMergedRegion(mergeRange);

				} catch (Exception e) {
					log.warn("Erro ao mesclar células - linha {}-{}, coluna {}", rowIndex, endRow, mapping.getIndex(),
							e);
				}
			}
		}
	}

	/**
	 * Valida se o exportador não foi fechado.
	 * 
	 * @throws IllegalStateException se o exportador estiver fechado
	 */
	private void validateNotClosed() {
		if (isClosed) {
			throw new IllegalStateException("Exportador já foi fechado");
		}
	}

	@Override
	public void close() throws IOException {
		if (isClosed) {
			return;
		}

		try {
			// O dispose() é crucial para SXSSF pois remove os arquivos XML temporários do
			// disco
			workbook.dispose();
			workbook.close();

			log.debug("Recursos do ExcelExport fechados com sucesso");

		} catch (Exception e) {
			log.error("Erro ao fechar recursos do ExcelExport", e);
			throw new IOException("Erro ao fechar recursos", e);

		} finally {
			isClosed = true;
		}
	}
}