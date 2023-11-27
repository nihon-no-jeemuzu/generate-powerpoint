package com.jemuzu.generate_powerpoint.utillity;

import java.io.FileReader;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;

import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;
import org.json.simple.parser.ContainerFactory;
import org.json.simple.parser.JSONParser;
import org.springframework.stereotype.Component;
import org.springframework.util.ObjectUtils;

import com.jemuzu.generate_powerpoint.model.PowerpointData;
import com.opencsv.CSVReader;

import lombok.extern.slf4j.Slf4j;

@Slf4j
@Component
public class ParseData {

	public PowerpointData parseSourceData(String sourcePath) throws Exception {
		AtomicInteger rowCounter = new AtomicInteger(0);
		List<List<String>> rowsData = new ArrayList<List<String>>();
		PowerpointData pptData = new PowerpointData();

		if (StringUtils.equals(FilenameUtils.getExtension(sourcePath), "csv")) {
			log.debug("Parsing csv data...");
			try (CSVReader reader = new CSVReader(new FileReader(sourcePath))) {
				String[] lineInArray;
				while ((lineInArray = reader.readNext()) != null) {
					if (rowCounter.getAndIncrement() == 0) {
						pptData.setHeaders(Arrays.asList(lineInArray));
					} else {
						rowsData.add(Arrays.asList(lineInArray));
					}
				}
				pptData.setRows(rowsData);
			}
		} else if (StringUtils.equals(FilenameUtils.getExtension(sourcePath), "json")) {
			log.debug("Parsing json data...");
			ContainerFactory containerFactory = new ContainerFactory() {
				@Override
				public Map createObjectContainer() {
					return new LinkedHashMap<>();
				}

				@Override
				public List creatArrayContainer() {
					return new LinkedList<>();
				}
			};

			List jsonMap = (List) new JSONParser().parse(new FileReader(sourcePath), containerFactory);
			List<String> headers = new ArrayList<String>();

			for (Object line : jsonMap) {
				List<String> rows = new ArrayList<String>();
				((Map) line).forEach((key, value) -> {
					headers.add(ObjectUtils.nullSafeToString(key));
					rows.add(ObjectUtils.nullSafeConciseToString(value));
				});
				rowsData.add(rows);
			}

			pptData.setHeaders(headers.stream().distinct().collect(Collectors.toList()));
			pptData.setRows(rowsData);
		}
		return pptData;
	}
}
