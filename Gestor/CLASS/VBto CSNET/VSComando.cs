using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Compatibility;

namespace VSClass
{
	public class VSComando
	{

		//=========================================================

		// VBto upgrade warning: Comando As adodb.Command	OnWrite(int)
		 private adodb.Command Comando = new adodb.Command();
		enum enuTipoComando {
			cmdArquivo = 256,
			cmdStoredProcedure = 4,
			cmdTabela = 2,
			cmdTabelaDireta = 512,
			cmdTexto = 1,
			cmdDesconhecido = 8
		};
		enum enuTipoDado {
			tipArray = 8192,
			tipBigInt = 20,
			tipBinary = 128,
			tipBoolean = 11,
			tipBSTR = 8,
			tipChapter = 136,
			tipChar = 129,
			tipCurrency = 6,
			tipDate = 7,
			tipDBDate = 133,
			tipDBTime = 134,
			tipDBTimeStamp = 135,
			tipDecimal = 14,
			tipDouble = 5,
			tipEmpty = 0,
			tipError = 10,
			tipFileTime = 64,
			tipGUID = 72,
			tipDispatch = 9,
			tipInteger = 3,
			tipUnknown = 13,
			tipLongVarBinary = 205,
			tipLongVarChar = 201,
			tipLongVarWChar = 203,
			tipNumeric = 131,
			tipPropVariant = 138,
			tipSingle = 4,
			tipSmallint = 2,
			tipTinyInt = 16,
			tipUnsignedBigInt = 21,
			tipUnsignedInt = 19,
			tipUnsignedSmallint = 18,
			tipUnsignedTinyInt = 17,
			tipUserDefined = 132,
			tipVarBinary = 204,
			tipVarChar = 200,
			tipVariant = 12,
			tipVarNumeric = 139,
			tipVarWChar = 202,
			tipWChar = 130
		};
		enum enuDirecaoParametro {
			parEntrada = 1,
			parEntradaSaida = 3,
			parSaida = 2,
			parValorRetorno = 4,
			parDesconhecido = 0
		};

		private VSComando() : base()
		{
			Comando = new adodb.Command();
		}
		~VSComando()
		{
			Comando = null;
		}


		public void Texto(object Bdados, string TSQL, enuTipoComando Tipo)
		{

			Comando.ActiveConnection = Bdados.Conexao.DBConnection;
			Comando.CommandText = TSQL;
			Comando.CommandType = Tipo;
		}


		public void Executa()
		{
			Comando.Execute();
		}


		public void setarParametro(string Nome, enuTipoDado Tipo, enuDirecaoParametro Direcao)
		{
			setarParametro(Nome, Tipo, Direcao, 0);
		}
		public void setarParametro(string Nome, enuTipoDado Tipo, enuDirecaoParametro Direcao, int Tamanho)
		{
			setarParametro(Nome, Tipo, Direcao, Tamanho, null);
		}
		public void setarParametro(string Nome, enuTipoDado Tipo, enuDirecaoParametro Direcao, int Tamanho, object Valor)
		{
			 adodb.Parameter Parametro = new adodb.Parameter();

			Parametro = Comando.CreateParameter(Nome, Tipo, Direcao, Tamanho, Valor);
			Comando.Parameters.Append(Parametro);
		}


		public void /* adodb.Parameter */ Parametro(object Indice)
		{
			Parametro = Comando.Parameters(Indice);
		}

	}
}