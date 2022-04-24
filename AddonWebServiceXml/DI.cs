using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace AddonWebServiceXml
{
    public class DI
    {
        public static SAPbobsCOM.Company oCompany;
        private static SAPbouiCOM.Application oApp;

        public static void Inicialize()
        {
            //DI.oCompany = new SAPbobsCOM.Company();
            //DI.oCompany.language = SAPbobsCOM.BoSuppLangs.ln_Portuguese_Br;

            SAPbouiCOM.SboGuiApi sboGuiApi = new SAPbouiCOM.SboGuiApi();
            sboGuiApi.Connect(System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(1)));
            DI.oApp = sboGuiApi.GetApplication();
            DI.oCompany = (SAPbobsCOM.Company)DI.oApp.Company.GetDICompany();

            // Criação da tabela de usuário de configuração do addon.

            DI.createTable("addon_seguradora_conf" , "Configuração addon seguradora" , SAPbobsCOM.BoUTBTableType.bott_NoObject);
            DI.createFieldsTable("@addon_seguradora_conf", "email_seguradora", "E-mail seguradora",  SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100);
            DI.createFieldsTable("@addon_seguradora_conf", "email_emitente", "E-mail emitente", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100);
            DI.createFieldsTable("@addon_seguradora_conf", "smtp", "SMTP", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10);
            DI.createFieldsTable("@addon_seguradora_conf", "mensagem", "Mensagem", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100);

            // Criação da tabela de usuário para armazenar o diretorio da filial

            DI.createTable("addon_seguradora_conf_filiais", "Configuração addon seguradora filiais", SAPbobsCOM.BoUTBTableType.bott_NoObject);
            DI.createFieldsTable("@addon_seguradora_conf_filiais", "filial", "Filial", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100);
            DI.createFieldsTable("@addon_seguradora_conf_filiais", "diretorio", "Diretorio", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100);

            // Criação da tabela de usuário para armazenar as placas do veiculo segurados
            // depois precisa modificar essa tabela para ser do tipo objeto.

            DI.createTable("addon_seguradora_veiculos", "Configuração addon seguradora veiculos", SAPbobsCOM.BoUTBTableType.bott_NoObject);
            DI.createFieldsTable("@addon_seguradora_veiculos", "placa", "Placa", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100);
            DI.createFieldsTable("@addon_seguradora_veiculos", "seguro", "Seguro", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100);


        }

        public static void createTable(string tableName, string tableDescription, SAPbobsCOM.BoUTBTableType type)
        {
            SAPbobsCOM.UserTablesMD oUserTable;
            oUserTable = (SAPbobsCOM.UserTablesMD)DI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
            bool ret = oUserTable.GetByKey(tableName);

            if (!ret)
            {
                // criação da tabela de configuração.
                oUserTable.TableName = tableName;
                oUserTable.TableDescription = tableDescription;
                oUserTable.TableType = type;
                oUserTable.Add();
                DI.oCompany.GetLastError(out int lErrCode, out string sErrMsg);
                if (lErrCode != 0)
                {
                    System.Windows.Forms.MessageBox.Show("Erro na criação da tabela de usuário: " + sErrMsg, "Erro", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
                    System.Environment.Exit(0);
                }

            }
        }

        public static void createFieldsTable(string tableName, string name, string description, SAPbobsCOM.BoFieldTypes type,
            SAPbobsCOM.BoFldSubTypes subType, int valueSize )
        {
            // criação das colunas na tabela criada
            SAPbobsCOM.UserFieldsMD oUserFields;
            oUserFields = (SAPbobsCOM.UserFieldsMD)(DI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields));

            oUserFields.TableName = tableName;
            oUserFields.Name = name;
            oUserFields.Description = description;
            oUserFields.Type = type;
            oUserFields.SubType = subType;
            oUserFields.EditSize = valueSize;

            int returnFields = oUserFields.Add();
            if (returnFields < 0)
            {
                System.Windows.Forms.MessageBox.Show("Houve um erro na criação do campo de usuário: ", "Erro", System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Error);
                System.Environment.Exit(0);
            }
        }
    }
}
