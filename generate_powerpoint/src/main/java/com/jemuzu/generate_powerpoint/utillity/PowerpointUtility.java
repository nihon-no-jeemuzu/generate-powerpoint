package com.jemuzu.generate_powerpoint.utillity;

import java.io.BufferedOutputStream;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URISyntaxException;
import java.nio.file.CopyOption;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Iterator;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;

import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.collections4.ListUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.Units;
import org.apache.poi.xddf.usermodel.chart.XDDFBarChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFAutoShape;
import org.apache.poi.xslf.usermodel.XSLFChart;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableRow;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.springframework.stereotype.Component;

import com.jemuzu.generate_powerpoint.model.PowerpointData;

import lombok.extern.slf4j.Slf4j;

@Slf4j
@Component
public class PowerpointUtility implements AppConstants {

	private PowerpointData pptData;
	private Path pptFilePath;
	private String targetPath;

	private XMLSlideShow pptTemplate;
	private XMLSlideShow ppt;
	//private Path pptFilePath;

	public PowerpointUtility() {
	}

	public PowerpointUtility(PowerpointData pptData, String targetPath) throws IOException, URISyntaxException {
		this.pptData = pptData;
		this.targetPath = targetPath;

		String templateDir = Paths.get(PowerpointUtility.class.getResource("/Template.pptx").toURI()).toFile().getAbsolutePath();
		pptFilePath = Files.copy(Paths.get(templateDir), Paths.get(this.targetPath, "generate_ppt_output.pptx"), StandardCopyOption.REPLACE_EXISTING);

		log.debug("template dir=" + templateDir);
		log.debug("ppt file in dir=" + pptFilePath);

		// Template
		FileInputStream templatePPT = new FileInputStream(templateDir);
		pptTemplate = new XMLSlideShow(templatePPT);
		templatePPT.close();

		// Report
		FileInputStream inputPPT = new FileInputStream(pptFilePath.toFile());
		ppt = new XMLSlideShow(inputPPT);
		pptTemplate.getSlides().forEach(s-> ppt.removeSlide(0));
		inputPPT.close();

	}

	public void export() throws FileNotFoundException, IOException {
		try (BufferedOutputStream out = new BufferedOutputStream(new FileOutputStream(pptFilePath.toFile()))) {
			log.debug("Creating file");
			writeData();
			ppt.write(out);
			out.close();
			log.debug("File Created");
		} catch (IOException e) {
			log.error("Error encountered while exporting PowerPoint.");
			throw e;
		} finally {
			log.debug("Closing PowerPoint documents...");
			if (ppt != null) {
				ppt.close();
			}
			if (pptTemplate != null) {
				pptTemplate.close();
			}
		}
	}

	private void writeData() {

		// cover
		log.debug("writing Cover Slide");
		XSLFSlide coverTemplate = pptTemplate.getSlides().get(PPT_COVER_SLIDE); // Cover Slide Template
		XSLFSlide coverSlide = ppt.createSlide(coverTemplate.getSlideLayout()).importContent(coverTemplate);
//		setXSLFTextBoxText(coverSlide, "", "");

		// charts
		log.debug("writing Charts Slide");
		XSLFSlide chartTemplate = pptTemplate.getSlides().get(PPT_CHART_SLIDE);
		XSLFSlide chartSlide = ppt.createSlide(chartTemplate.getSlideLayout()).importContent(chartTemplate);
		XSLFChart barChart = findCharts(chartSlide).get(0);
		
		String[] series = pptData.getHeaders().stream().toArray(String[]::new);
		String[] categories = pptData.getRows().stream().map(c -> c.get(0)).toArray(String[]::new);
	
		AtomicInteger counter = new AtomicInteger(1);
		List<Double[]> values = new ArrayList<Double[]>() {
			{
				while (counter.get() < Arrays.asList(series).size()) {
					add(pptData.getRows().stream()
							.map(r -> r.get(counter.get()))
							.map(s -> s.replaceAll(",", ""))
							.map(Double::parseDouble)
							.toArray(Double[]::new));
					counter.getAndIncrement();
				}
			}
		};
		populateBarChart(barChart, series, categories, values);

		// tables
		log.debug("writing Table Slides");
		XSLFSlide tableTemplate = pptTemplate.getSlides().get(PPT_TABLE_SLIDE);
		
		ListUtils.partition(pptData.getRows(), 20).forEach(subList -> {
			XSLFSlide tableSlide = ppt.createSlide(tableTemplate.getSlideLayout()).importContent(tableTemplate);
			XSLFTable table = findTables(tableSlide).get(0);
			AtomicInteger column_counter = new AtomicInteger(0);

			XSLFTableRow headerRow = table.getRows().get(0);
			headerRow.setHeight(0.5);
			pptData.getHeaders().forEach(header -> {
				if(column_counter.get() == 0) {
					XSLFTextRun tr = headerRow.getCells().get(0)
						.getTextParagraphs().get(0)
						.getTextRuns().get(0);
					tr.setFontSize(12.0);
					tr.setText(header);
					table.setColumnWidth(column_counter.get(), Units.POINT_DPI * 4);
				} else {
					XSLFTextRun tr =headerRow.addCell()
						.addNewTextParagraph()
						.addNewTextRun();
					tr.setFontSize(11.0);
					tr.setText(header);
					table.setColumnWidth(column_counter.get(), Units.POINT_DPI * 0.95);
				}
				column_counter.getAndIncrement();
			});
			
			subList.forEach(data -> {
				XSLFTableRow row = table.addRow();
				row.setHeight(0.5);
				data.forEach(value -> {
					XSLFTextRun tr = row.addCell().addNewTextParagraph().addNewTextRun();
					tr.setFontSize(11.0);
					tr.setText(value);
				});
			});
		});
	}

	/*
	 * UTILITIES
	 */

	private void setXSLFTextBoxText(XSLFSlide slide, String text, String shapeName) {
		slide.getShapes().stream()
			.filter(shape -> StringUtils.equals(shape.getShapeName(), shapeName))
			.findFirst().ifPresent(shapeObj -> {
				if (shapeObj instanceof XSLFTextShape) {
					XSLFTextShape textShape = (XSLFTextShape) shapeObj;
					textShape.setText(text);
				} else if (shapeObj instanceof XSLFAutoShape) {
					XSLFAutoShape autoShape = (XSLFAutoShape) shapeObj;
					autoShape.setText(text);
				} else {
					XSLFTextBox tbShape = (XSLFTextBox) shapeObj;
					tbShape.setText(text);
				}
			});
	}
	
	private void populateBarChart(XSLFChart chart, String[] series, String[] categories, List<Double[]> values) {
		List<XDDFChartData> data = chart.getChartSeries();
		XDDFBarChartData bar = (XDDFBarChartData) data.get(0);
		int numOfPoints = categories.length;

		if (numOfPoints != 0) {
			String categoryDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, 0, 0));
			XDDFDataSource<?> categoriesData = XDDFDataSourcesFactory.fromArray(categories, categoryDataRange, 0);

			// if series in the template is more than the values
			while (values.size() < bar.getSeriesCount()) {
				bar.removeSeries(values.size());
			}

			AtomicInteger index = new AtomicInteger(1);
			values.forEach(value -> {
				String valuesDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, index.get(), index.get()));
				XDDFNumericalDataSource<? extends Number> valuesData = XDDFDataSourcesFactory.fromArray(value, valuesDataRange, index.get());
				valuesData.setFormatCode("General");
				XDDFChartData.Series series1 = null;

				// if values is more than the template series
				if (index.get() <= bar.getSeriesCount()) {
					series1 = bar.getSeries(index.get() - 1);
					series1.replaceData(categoriesData, valuesData);
				} else {
					series1 = bar.addSeries(categoriesData, valuesData);
				}

				series1.setTitle(series[index.get()], chart.setSheetTitle(series[index.get()], index.get()));
				index.getAndIncrement();
			});
			chart.plot(bar);
		}
	}

	private List<XSLFChart> findCharts(XSLFSlide slide) {
		// find all charts in the slide
		List<XSLFChart> chart = new ArrayList<XSLFChart>();
		chart.addAll(slide.getRelations().stream().filter(XSLFChart.class::isInstance).map(XSLFChart.class::cast).collect(Collectors.toList()));
		if (chart.size() == 0) {throw new IllegalStateException("chart not found in the template");}
		return chart;
	}
	
	private List<XSLFTable> findTables(XSLFSlide slide) {
		// find all tables in the slide
		List<XSLFTable> table = new ArrayList<XSLFTable>();
		table.addAll(slide.getShapes().stream().filter(XSLFTable.class::isInstance).map(XSLFTable.class::cast).collect(Collectors.toList()));
		if (table.size() == 0) {throw new IllegalStateException("table not found in the template");}
		return table;
	}

}
