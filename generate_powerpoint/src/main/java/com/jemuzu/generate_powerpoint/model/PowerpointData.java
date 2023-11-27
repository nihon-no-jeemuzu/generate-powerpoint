package com.jemuzu.generate_powerpoint.model;

import java.util.List;

import lombok.Data;

@Data
public class PowerpointData {

	private List<String> headers;
	private List<List<String>> rows;

}
