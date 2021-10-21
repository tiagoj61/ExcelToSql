public class Main {

	public static void main(String[] args)
			throws EncryptedDocumentException, InvalidFormatException, FileNotFoundException, IOException {

		File pasta = new File("C:\\Users\\Tiago\\Downloads\\teste");
		for (File arquivo : pasta.listFiles()) {
			File comandos = new File("C:\\Users\\Tiago\\Downloads\\testecomands\\" + arquivo.getName() + ".txt");
			if (!comandos.exists())
				comandos.delete();
			comandos.createNewFile();
			FileWriter escreve = new FileWriter(comandos);
			BufferedWriter esc = new BufferedWriter(escreve);
			XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(arquivo));

			XSSFSheet sheet = workbook.getSheetAt(0);
			DataFormatter dataFormatter = new DataFormatter();

			Iterator<Row> rowIterator = sheet.rowIterator();
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();

				Iterator<Cell> cellIterator = row.cellIterator();
				String nomePasta = "";
				int pastasPlavrasChaves = 0;
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					String cellValue = dataFormatter.formatCellValue(cell);
					if (pastasPlavrasChaves < 2) {
						if (!cellValue.isEmpty()) {
							nomePasta += cellValue + " ";
						}
					} else {
						if (pastasPlavrasChaves == 2) {
							String DMLPasta = "INSERT INTO pasta(ativo, nome) VALUES(TRUE,";
							DMLPasta += "'" + nomePasta + "'";
							DMLPasta += ");";
//
							if (nomePasta != "" && nomePasta != " " && nomePasta.matches(".*\\d.*")) {
								esc.append(DMLPasta);
								esc.newLine();
							}
						}
						if (nomePasta != "" && nomePasta != " " && nomePasta.matches(".*\\d.*")) {
							String palavraChave = "";
							palavraChave += cellValue;

							String DMLPalavraChave = "INSERT INTO public.palavrachave(nome,tipodapalavraenum,pasta_id)";
							DMLPalavraChave += "(SELECT '" + palavraChave
									+ "',0,pasta.id FROM pasta ORDER BY id DESC limit 1);";
							esc.append(DMLPalavraChave);
							esc.newLine();

							if (palavraChave.contains(".") || palavraChave.contains("-")) {
								palavraChave = palavraChave.replace(".", "");
								palavraChave = palavraChave.replace("-", "");
								palavraChave = palavraChave.replace(",", "");
								DMLPalavraChave = "INSERT INTO public.palavrachave(nome,tipodapalavraenum,pasta_id)";
								DMLPalavraChave += "(SELECT '" + palavraChave
										+ "',0,pasta.id FROM pasta ORDER BY id DESC limit 1);";
								esc.append(DMLPalavraChave);
								esc.newLine();
							}
						}
					}
					pastasPlavrasChaves++;
				}
				if (nomePasta != "" && nomePasta != " " && nomePasta.matches(".*\\d.*")) {
					String NV2 = arquivo.getName().split("N2")[1].split("-")[0].trim();
					String DMLGrupoPasta = "INSERT INTO public.grupo_pasta(grupo_id,pasta_id)";
					DMLGrupoPasta += "(SELECT '" + NV2 + "',pasta.id FROM pasta ORDER BY id DESC limit 1);";
					esc.append(DMLGrupoPasta);
					esc.newLine();
				}

				// generateCreatePasta(nomePasta, comandos);
			}
			esc.close();
		}

	}
}