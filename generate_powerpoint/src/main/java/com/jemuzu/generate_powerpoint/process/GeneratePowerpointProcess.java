package com.jemuzu.generate_powerpoint.process;

import java.nio.file.Files;
import java.nio.file.LinkOption;
import java.nio.file.Paths;

import org.apache.commons.collections4.ListUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import com.jemuzu.generate_powerpoint.model.PowerpointData;
import com.jemuzu.generate_powerpoint.utillity.ParseData;
import com.jemuzu.generate_powerpoint.utillity.PowerpointUtility;

import lombok.extern.slf4j.Slf4j;

@Slf4j
@Component
public class GeneratePowerpointProcess {

	@Autowired
	ParseData parseData;

	public void executeProcess(String sourcePath, String targetPath) throws Exception {

		if (Files.isDirectory(Paths.get(targetPath), LinkOption.NOFOLLOW_LINKS)) {
			PowerpointData pptData = parseData.parseSourceData(sourcePath);

			// validate headers count
			if (ListUtils.emptyIfNull(pptData.getHeaders()).size() <= 10) {
				PowerpointUtility pptUtility = new PowerpointUtility(pptData, targetPath);
				pptUtility.export();
				log.info("Process Complete");
			} else {
				log.error("ERROR: Cannot process data with more than 10 headers."); // This limit is for aesthetic purposes only.
			}
		} else {
			log.error("ERROR: targetPath does not exists.");
			System.exit(1);
		}
	}
}
