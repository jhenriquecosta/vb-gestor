package convenio;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;

public class Retorno {
	static Connection cnn;
	public static void main(String[] args) {
	      	try{
	      		cnn = getCPTRIBConnection();
	      		
	      		//File dir = new File("c:/impostodigital/basa/convenio/data");
	      		File dir = new File("c:/impostodigital/bb");
	    		
	      		File[] files = dir.listFiles();
	    		String bancoNossoNumero;
	    		String prefeituraNumeroDocumento;
	    		String status;
	    		String legenda;
	    		String dataCredito;
	    		String valorBoleto;
	    		for (File file: files){
	    			//if(file.getName().startsWith("CBR643")){
	    				try {
	    					FileReader reader = new FileReader(file);
	    					BufferedReader bf = new BufferedReader(reader);
	    					String line="";
	    					System.out.println("Documento --- Nosso Numero --- Status --- Data Credito --- Valor");
	    					while((line = bf.readLine())!=null){
	    						if(!line.contains("COBRANCA") && !line.startsWith("9")){
	    							//bancoNossoNumero = basaNossoNumero(line);
	    							prefeituraNumeroDocumento = bbNumeroDocumento(line);
	    							updateObrigacao(prefeituraNumeroDocumento);
	    							
	    							
	    							//status = basaNumeroOcorrencia(line);
	    							//legenda= basaLegenda(status);
	    							//dataCredito = basaDataCredito(line);
	    							//valorBoleto = basaValor(line);
	    							//insertCPTRIB(prefeituraNumeroDocumento,bancoNossoNumero,  status, legenda, dataCredito, valorBoleto);
	    							System.out.println(prefeituraNumeroDocumento + " --- " + file.getName());
	    						}
	    					}
	    					bf.close();
	    					reader.close();
	    				} catch (IOException e) {
	    					e.printStackTrace();
	    				}
	    			}
	    			
	    		//}
	      		cnn.close();
	      		System.out.println("FIM");
	      	}catch(Exception ex){
	      		ex.printStackTrace();
	      	}	
	}  
	static String formatCredito(String dataCredito){
		if(dataCredito!=null && dataCredito.length() >=6){
			String dc = dataCredito.substring(0,2) + "/" + dataCredito.substring(2,4) + "/" + dataCredito.substring(4,6) ;
			return dc;
		}else
			return "";
	}
	static void updateObrigacao(String documento)throws SQLException,ClassNotFoundException{
		Long numero = Long.valueOf(documento);
		String sql = "UPDATE TAB_OBRIGACAO_CONTRIBUINTE SET TOC_STATUS_OBRIGACAO = 3 WHERE TOC_COD_OBRIGACAO =  ?";
		PreparedStatement pst=cnn.prepareStatement(sql);
		pst.setString(1, numero.toString());
		pst.executeUpdate();
	}
	static void insertCPTRIB(String documento,String nossoNumero, String status, String legenda,String dataCredito,String valor)throws SQLException,ClassNotFoundException{
		String sql = "INSERT INTO TAB_BCP_RETORNO_GERAL (DOCUMENTO,NOSSO_NUMERO,STATUS,LEGENDA,DATA_CREDITO,VALOR) VALUES (?,?,?,?,?,?)";
		PreparedStatement pst=cnn.prepareStatement(sql);
		pst.setObject(1, new Long(documento).toString());
		pst.setObject(2, nossoNumero);
		pst.setObject(3, status);
		pst.setObject(4, legenda);
		pst.setObject(5, dataCredito);
		pst.setObject(6, valor);
		pst.executeUpdate();
		pst.close();
	}
	static Connection getCPTRIBConnection() throws SQLException,ClassNotFoundException{
		String connectionUrl = "jdbc:sqlserver://192.168.1.254:1433;databaseName=CPTRIB;user=sa;password=kabecao";
		// Establish the connection.  
        Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");  
        return DriverManager.getConnection(connectionUrl);  
	}
	static String basaNumeroOcorrencia(String line){
		return  line.substring(109-1, 110);
	}
	static String basaDataOcorrencia(String line){
		return  line.substring(111-1, 116);
	}
	static String basaNumeroDocumento(String line){
		return  line.substring(117-1, 126);
	}
	static String basaValor(String line){
		return  line.substring(153-1, 165);
	}
	static String basaDataCredito(String line){
		return  line.substring(296-1, 301);
	}
	static String basaNossoNumero(String line){
		return  line.substring(71-1, 82);
	}
	static String bbNumeroDocumento(String line){
		return  line.substring(71-1, 80);
	}
	static String bbDataCredito(String line){
		return  line.substring(176-1, 181);
	}
	static String bbStatus(String line){
		return  line.substring(109-1, 110);
	}
	static String basaLegenda(String status){
		String legenda = "Não indentificado";
		switch (status) {
		case "02":
			legenda = "Entrada Confirmada";
			break;
		case "03":
			legenda = "Rejeitada";
			break;
		case "06":
			legenda = "Liquidação Normal";
			break;
		case "09":
			legenda = "Liquidação Normal";
			break;
		case "10":
			legenda = "Liquidação Normal";
			break;
		}
		return legenda;
		
	}
	static String bbLegenda(String status){
		String legenda = "liquidação normal";
		switch (status) {
		case "02":
			legenda = "liquidação parcial";
			break;
		case "03":
			legenda = "liquidação por saldo";
			break;
		case "04":
			legenda = "liquidação com cheque a compensar";
			break;
		case "05":
			legenda = "liquidação de título sem registro (carteira 7 tipo 4)";
			break;
		case "07":
			legenda = "liquidação na apresentação";
			break;
		case "09":
			legenda = "liquidação em cartório";
			break;
		}
		return legenda;
		
	}
}
