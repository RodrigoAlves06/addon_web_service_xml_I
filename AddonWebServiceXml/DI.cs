using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM;


namespace AddonWebServiceXml
{
    public class DI
    {
        public static SAPbobsCOM.Company oCompany;
        public static SAPbouiCOM.Application oApp;

        public static void Inicialize()
        {
            //DI.oCompany = new SAPbobsCOM.Company();
            //DI.oCompany.language = SAPbobsCOM.BoSuppLangs.ln_Portuguese_Br;

            SAPbouiCOM.SboGuiApi sboGuiApi = new SAPbouiCOM.SboGuiApi();
            sboGuiApi.Connect(System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(1)));
            DI.oApp = sboGuiApi.GetApplication();
            DI.oCompany = (SAPbobsCOM.Company)DI.oApp.Company.GetDICompany();

            DI.oApp.SetStatusBarMessage(string.Format("Addon - Verificação de tabelas e campos de usuários...",
                    System.Windows.Forms.Application.ProductName),
                    BoMessageTime.bmt_Medium, false);

            // Criação da tabela de usuário de configuração do addon.

            bool returnCreateTable = true;

            returnCreateTable =  DI.createTable("seguradora_conf" , "Configuração addon seguradora" , SAPbobsCOM.BoUTBTableType.bott_NoObject);
            if(returnCreateTable == false)
            {
                DI.createFieldsTable("seguradora_conf", "email_seguradora", "E-mail seguradora",
                    SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
                DI.createFieldsTable("seguradora_conf", "email_emitente", "E-mail emitente", SAPbobsCOM.BoFieldTypes.db_Alpha,
                     100);
                DI.createFieldsTable("seguradora_conf", "smtp", "SMTP", SAPbobsCOM.BoFieldTypes.db_Alpha,
                     10);
                DI.createFieldsTable("seguradora_conf", "mensagem", "Mensagem", SAPbobsCOM.BoFieldTypes.db_Alpha,
                     100);
                DI.createFieldsTable("seguradora_conf", "assunto", "Assunto", SAPbobsCOM.BoFieldTypes.db_Alpha,
                    100);
                DI.createFieldsTable("seguradora_conf", "pass", "Password", SAPbobsCOM.BoFieldTypes.db_Alpha,
                    100);

            }


            // Criação da tabela de usuário para armazenar o diretorio da filial

            returnCreateTable = DI.createTable("seg_conf_filiais", "Addon seguradora filiais", SAPbobsCOM.BoUTBTableType.bott_NoObject);
            if (returnCreateTable == false)
            {
                DI.createFieldsTable("seg_conf_filiais", "filial", "Filial", SAPbobsCOM.BoFieldTypes.db_Alpha,
                    100);
                DI.createFieldsTable("seg_conf_filiais", "diretorio", "Diretorio", SAPbobsCOM.BoFieldTypes.db_Alpha,
                     100);

            }


            // Criação da tabela de usuário para armazenar as placas do veiculo segurados
            // depois precisa modificar essa tabela para ser do tipo objeto.

            returnCreateTable = DI.createTable("seg_veiculos", "Addon seguradora veiculos", SAPbobsCOM.BoUTBTableType.bott_NoObject);
            if (returnCreateTable == false)
            {
                DI.createFieldsTable("@seg_veiculos", "placa", "Placa", SAPbobsCOM.BoFieldTypes.db_Alpha,
                100);
                DI.createFieldsTable("@seg_veiculos", "seguro", "Seguro", SAPbobsCOM.BoFieldTypes.db_Alpha,
                     100);
            }


        }

        public static bool createTable(string tableName, string tableDescription, SAPbobsCOM.BoUTBTableType type)
        {
            SAPbouiCOM.SboGuiApi sboGuiApi = new SAPbouiCOM.SboGuiApi();
            sboGuiApi.Connect(System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(1)));
            DI.oApp = sboGuiApi.GetApplication();
            DI.oCompany = (SAPbobsCOM.Company)DI.oApp.Company.GetDICompany();

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
                    //System.Windows.Forms.MessageBox.Show("Erro na criação da tabela de usuário: " + sErrMsg, "Erro", System.Windows.Forms.MessageBoxButtons.OK,
                    //System.Windows.Forms.MessageBoxIcon.Error);
                    //System.Environment.Exit(0);
                }

            }
            return ret;
        }

        //SAPbobsCOM.BoFldSubTypes subType

        public static void createFieldsTable(string tableName, string name, string description, SAPbobsCOM.BoFieldTypes type,int valueSize )
        {
            SAPbouiCOM.SboGuiApi sboGuiApi = new SAPbouiCOM.SboGuiApi();
            sboGuiApi.Connect(System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(1)));
            DI.oApp = sboGuiApi.GetApplication();
            DI.oCompany = (SAPbobsCOM.Company)DI.oApp.Company.GetDICompany();

            // criação das colunas na tabela criada
            SAPbobsCOM.UserFieldsMD oUserFieldsMD;
            oUserFieldsMD = (SAPbobsCOM.UserFieldsMD)(DI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields));

            oUserFieldsMD.TableName = tableName;
            oUserFieldsMD.Name = name;
            oUserFieldsMD.Description = description;
            oUserFieldsMD.Type = type;
            //oUserFieldsMD.SubType = subType;
            oUserFieldsMD.EditSize = Convert.ToInt32(valueSize);

            int returnFields = oUserFieldsMD.Add();
            //DI.oCompany.GetLastError(out int lErrCode, out string sErrMsg);

            if (returnFields < 0)
            {
                //System.Windows.Forms.MessageBox.Show("Houve um erro na criação do campo de usuário: " , "Erro", System.Windows.Forms.MessageBoxButtons.OK,
                //System.Windows.Forms.MessageBoxIcon.Error);
                //System.Environment.Exit(0);
                Console.WriteLine("erro");
            }
        }
    }
}
