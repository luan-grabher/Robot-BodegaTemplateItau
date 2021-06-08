package ContabilityTemplateImportation;

import Entity.Executavel;
import JExcel.XLSX;
import TemplateContabil.Model.Entity.Importation;
import TemplateContabil.Model.ImportationModel;
import fileManager.FileManager;
import java.io.File;
import java.math.BigDecimal;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Model {

    public class converterArquivoParaTemplate extends Executavel {

        private final Importation importation;
        private final Integer mes;
        private final Integer ano;

        public converterArquivoParaTemplate(Importation importation, Integer mes, Integer ano) {
            this.importation = importation;
            this.mes = mes;
            this.ano = ano;
        }

        @Override
        public void run() {
            //Pega o arquivo PDF da importation
            File file = importation.getFile();

            //Cria Configuração
            Map<String, Map<String, String>> cfgCols = new HashMap<>();
            cfgCols.put("data", XLSX.convertStringToConfig("data", "-collumn¬A¬-type¬string¬-required¬true¬-regex¬[0-9]{2} *\\/ *[a-zA-z]{3}"));
            cfgCols.put("hist", XLSX.convertStringToConfig("hist", "-collumn¬B¬-type¬string¬-required¬true"));
            cfgCols.put("agencia", XLSX.convertStringToConfig("agencia", "-collumn¬C¬-type¬string"));
            cfgCols.put("valor", XLSX.convertStringToConfig("valor", "-collumn¬D§E¬-type¬string¬-required¬true"));

            //Pega os dados do Excel
            List<Map<String, Object>> rows = XLSX.get(file, cfgCols);

            StringBuilder csvtext = new StringBuilder("#data;historico;valor");

            //Percorre Excel
            for (Map<String, Object> row : rows) {
                if (!row.get("hist").toString().contains("SALDO")) {
                    //Arruma data deixando apenas o dia
                    String data = row.get("data").toString().replaceAll("[^0-9]*", "");
                    data += "/" + mes + "/" + ano;

                    csvtext.append("\r\n");
                    csvtext.append(data).append(";");

                    csvtext.append(row.get("hist").toString());
                    //Se tiver agencia
                    if (row.containsKey("agencia") && row.get("agencia") != null) {
                        csvtext.append(" Agencia Origem: ").append(row.get("agencia").toString());
                    }
                    csvtext.append(";");                   

                    csvtext.append(row.get("valor").toString().trim());
                }
            }

            //Salva arquivo como CSV
            File newFile = new File(file.getParent() + "\\" + file.getName().replaceAll(".xlsx", ".csv"));
            FileManager.save(newFile, csvtext.toString());

            //Troca o arquivo file da importation
            importation.setFile(newFile);

            //Chama o modelo da importação que irá criar o template e gerar warning se algo der errado
            ImportationModel modelo = new ImportationModel(importation.getNome(), mes, ano, importation, null);

            //Pega lctos
            //List<LctoTemplate> lctos = importation.getLctos();
            modelo.criarTemplateDosLancamentos(importation);
        }
    }
}
