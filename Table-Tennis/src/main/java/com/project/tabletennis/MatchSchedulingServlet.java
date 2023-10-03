package com.project.tabletennis;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.URI;
import java.net.URL;
import java.nio.file.*;
import javax.servlet.*;
import javax.servlet.annotation.*;
import javax.servlet.http.*;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collections;
import java.util.List;

@WebServlet("/MatchSchedulingServlet")
public class MatchSchedulingServlet extends HttpServlet {
	private static final long serialVersionUID = 1L;

	protected void doPost(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {
		response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
		response.setHeader("Content-Disposition", "attachment; filename=match_schedule.xlsx");

		try (OutputStream outputStream = response.getOutputStream()) {
			// Download the CSV file from a URL
			String fileURL = "https://drive.google.com/uc?export=download&id=1YEcmKlGULxCnzE0TdflKYfplnuSyzznZ";
			String destinationPath = "F:\\Internshala task/csv_player_Registration.csv";

			try {
				URI uri = new URI(fileURL);
				URL url = uri.toURL();
				HttpURLConnection connection = (HttpURLConnection) url.openConnection();
				connection.setRequestMethod("GET");

				InputStream in = connection.getInputStream();
				Files.copy(in, Paths.get(destinationPath), StandardCopyOption.REPLACE_EXISTING);
				in.close();

				// Define players list
				List<String> players = new ArrayList<>();

				// Now, you can read and process the downloaded CSV file from 'destinationPath'
				// You can use a CSV parsing library as shown in the previous responses.
				try (CSVParser csvParser = new CSVParser(new FileReader(destinationPath), CSVFormat.DEFAULT)) {
					// Skip the header row
					boolean skipHeader = true;

					for (CSVRecord csvRecord : csvParser) {
						if (skipHeader) {
							skipHeader = false;
							continue; // Skip the header row
						}

						String playerName = csvRecord.get(1); // Assuming the player name is in the second column (index
																// 1)
						players.add(playerName); // Add player name to the list
					}
				}

				// Create an Excel workbook and write it to the response
				Workbook workbook = new XSSFWorkbook();
				Sheet sheet = workbook.createSheet("Match Schedule");

				// Create the header row
				Row headerRow = sheet.createRow(0);
				headerRow.createCell(0).setCellValue("Date");
				headerRow.createCell(1).setCellValue("Player 1");
				headerRow.createCell(2).setCellValue("Player 2");
				headerRow.createCell(3).setCellValue("Referee");

				// Create Calendar instance for match date
				Calendar matchDate = Calendar.getInstance();
				matchDate.set(2023, Calendar.OCTOBER, 3);

				// ... Populate match schedule data ...
				int rowNum = 1; // Start from the second row (after the header)
				for (int i = 0; i < players.size(); i += 2) {
					Row row = sheet.createRow(rowNum++);
					row.createCell(0).setCellValue(matchDate.getTime().toString());
					row.createCell(1).setCellValue(players.get(i));
					row.createCell(2).setCellValue(players.get(i + 1));
					row.createCell(3).setCellValue(
							(matchDate.get(Calendar.DAY_OF_WEEK) == Calendar.SUNDAY) ? "Referee_2" : "Referee_1");

					// Increment the match date (excluding Sundays)
					if (matchDate.get(Calendar.DAY_OF_WEEK) != Calendar.SUNDAY) {
						matchDate.add(Calendar.DATE, 1);
					}
				}

				// Write the workbook to the response output stream
				workbook.write(outputStream);

				// Close the workbook and output stream
				workbook.close();
				outputStream.close();

			} catch (Exception e) {
				e.printStackTrace();
				response.getWriter().println("Error downloading or processing the CSV file.");
			}
		}
	}
}
