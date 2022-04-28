using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using SAPbouiCOM;
using SAPbobsCOM;
using System.Net.Mail;

namespace AddonWebServiceXml
{
    public class FormWebServiceXml
    {
        private SAPbouiCOM.Application oApplication;
        private SAPbouiCOM.Form oForm;
        private SAPbouiCOM.Item oItem;
        private SAPbouiCOM.Grid oGrid;
        private static SAPbobsCOM.Company _company;
        private SAPbobsCOM.Recordset oRecordset;
        private static SAPbouiCOM.Application oApp;
           

        private void SetApplication()
        {
            SboGuiApi oSboGuiApi = null;
            string sConnectionString = null;

            oSboGuiApi = new SAPbouiCOM.SboGuiApi();

            sConnectionString = Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));

            try
            {
                oSboGuiApi.Connect(sConnectionString);
                oApp = oSboGuiApi.GetApplication();
                _company = (SAPbobsCOM.Company)oApp.Company.GetDICompany();
            }
            catch(Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Erro de conexão Addon alfa seguradora: " + ex.Message);
                System.Environment.Exit(0);
            }

            oApplication = oSboGuiApi.GetApplication(-1);
            oApplication.SetStatusBarMessage(string.Format("Addon alfa seguradora conectado com sucesso!",
                System.Windows.Forms.Application.ProductName),
                BoMessageTime.bmt_Medium, false);
        }

        private void createForm()
        {
            try
            {
                SAPbouiCOM.Button oButton = null;
                SAPbouiCOM.Folder oFolder = null;
                SAPbouiCOM.ComboBox oComboBox = null;
                SAPbouiCOM.StaticText oStaticText = null;
                SAPbouiCOM.EditText oEditText = null;
                SAPbouiCOM.OptionBtn oPtionBtn = null;
                SAPbouiCOM.FormCreationParams oCreationParams = null;
                oCreationParams = ((SAPbouiCOM.FormCreationParams)(oApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));

                oCreationParams.UniqueID = "frmWebServXml";
                oCreationParams.FormType = "frmWebServXml";
                oCreationParams.BorderStyle = BoFormBorderStyle.fbs_Sizable;

                oForm = oApplication.Forms.AddEx(oCreationParams);


                AddDataSourceNoForm();

                oForm.Title = "Addon - Alfa seguradora";
                oForm.Left = 300;
                //oForm.ClientWidth = 200;
                //oForm.Top = 100;
                //oForm.ClientHeight = 140;

                oForm.ClientWidth = 1000;
                oForm.Top = 100;
                oForm.ClientHeight = 640;

                oItem = oForm.Items.Add("btnSend", BoFormItemTypes.it_BUTTON);
                oItem.Left = 5;
                oItem.Width = 65;
                oItem.Top = 610;
                oItem.Height = 19;
                oButton = ((SAPbouiCOM.Button)(oItem.Specific));
                oButton.Caption = "Enviar XML";

                oItem = oForm.Items.Add("2", BoFormItemTypes.it_BUTTON);
                oItem.Left = 75;
                oItem.Width = 65;
                oItem.Top = 610;
                oItem.Height = 19;
                oButton = ((SAPbouiCOM.Button)(oItem.Specific));
                oButton.Caption = "Cancelar";

                // adicionando retangulo
                oItem = oForm.Items.Add("React", BoFormItemTypes.it_RECTANGLE);
                oItem.Left = 0;
                oItem.Width = 994;
                oItem.Top = 22;
                oItem.Height = 580;

                // adicionando folder I
                oItem = oForm.Items.Add("Conteudo", BoFormItemTypes.it_FOLDER);
                oItem.Left = 0;
                oItem.Width = 120;
                oItem.Top = 6;
                oItem.Height = 19;

                oFolder = ((SAPbouiCOM.Folder)(oItem.Specific));
                oFolder.Caption = "Conteúdo Envio XML";
                oFolder.DataBind.SetBound(true, "", "FolderDS");
                oFolder.Select();

                // adicionando folder II
                oItem = oForm.Items.Add("Conf", BoFormItemTypes.it_FOLDER);
                oItem.Left = 100;
                oItem.Width = 120;
                oItem.Top = 6;
                oItem.Height = 19;

                oFolder = ((SAPbouiCOM.Folder)(oItem.Specific));
                oFolder.Caption = "Configurações";
                oFolder.DataBind.SetBound(true, "", "FolderDS");
                oFolder.GroupWith("Conteudo"); // server para informar que deve ser agrupado ao lado de conteudo.

                // A


                // Adicionando Label
                oItem = oForm.Items.Add("lbl", BoFormItemTypes.it_STATIC);
                oItem.Left = 52;
                oItem.Width = 100;
                oItem.Top = 49;
                oItem.Height = 19;

                // comando para deixar visivel nesse folder
                oItem.FromPane = 1;
                oItem.ToPane = 1;

                oItem.LinkTo = "cmb"; //Serve para linkar o objeto com outro.

                oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
                oStaticText.Caption = "Filial:";

                // Adicionando combobox

                oItem = oForm.Items.Add("cmb", BoFormItemTypes.it_COMBO_BOX);
                oItem.Left = 90;
                oItem.Width = 330;
                oItem.Top = 49;
                oItem.Height = 19;

                oItem.DisplayDesc = true; // significa que quando selecionado a opção, o mesmo vai apresentar o ID do selecionado (false)
                                           // como true ele vai trazer o nome do valor;

                // comando para deixar visivel nesse folder
                oItem.FromPane = 1;
                oItem.ToPane = 1;

                // oItem.LinkTo = "EditText1"; Serve para linkar o objeto com outro.

                oComboBox = ((SAPbouiCOM.ComboBox)(oItem.Specific));

                //oStaticText.GroupWith("lblFilial");
                oComboBox.DataBind.SetBound(true, "", "ComboBoxDS");



                // Adicionando Label
                oItem = oForm.Items.Add("lblP", BoFormItemTypes.it_STATIC);
                oItem.Left = 450;
                oItem.Width = 120;
                oItem.Top = 49;
                oItem.Height = 19;

                // comando para deixar visivel nesse folder
                oItem.FromPane = 1;
                oItem.ToPane = 1;

                oItem.LinkTo = "lblDtI"; //Serve para linkar o objeto com outro.

                oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
                oStaticText.Caption = "Período de Faturamento";

                // Adicionando Label
                oItem = oForm.Items.Add("lblDtI", BoFormItemTypes.it_STATIC);
                oItem.Left = 570;
                oItem.Width = 120;
                oItem.Top = 49;
                oItem.Height = 19;

                // comando para deixar visivel nesse folder
                oItem.FromPane = 1;
                oItem.ToPane = 1;

                oItem.LinkTo = "txtDtI"; //Serve para linkar o objeto com outro.

                oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
                oStaticText.Caption = "Data Início:";

                // Adicionando TextBox: oEditText -- DataIni

                oItem = oForm.Items.Add("txtDtI", BoFormItemTypes.it_EXTEDIT);
                oItem.Left = 630;
                oItem.Width = 65;
                oItem.Top = 49;
                oItem.Height = 14;
                oItem.Enabled = true;
                // comando para deixar visivel nesse folder
                oItem.FromPane = 1;
                oItem.ToPane = 1;

                // oItem.LinkTo = "EditText1"; Serve para linkar o objeto com outro.

                oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));

                oEditText.DataBind.SetBound(true, "", "EditTextI");


                // Adicionando Label
                oItem = oForm.Items.Add("lblDtF", BoFormItemTypes.it_STATIC);
                oItem.Left = 570;
                oItem.Width = 120;
                oItem.Top = 69;
                oItem.Height = 19;

                // comando para deixar visivel nesse folder
                oItem.FromPane = 1;
                oItem.ToPane = 1;

                oItem.LinkTo = "txtDtF"; //Serve para linkar o objeto com outro.

                oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
                oStaticText.Caption = "Data Fim:";

                // Adicionando TextBox: oEditText DataFim

                oItem = oForm.Items.Add("txtDtF", BoFormItemTypes.it_EXTEDIT);
                oItem.Left = 630;
                oItem.Width = 65;
                oItem.Top = 69;
                oItem.Height = 14;
                oItem.Enabled = true;

                // comando para deixar visivel nesse folder
                oItem.FromPane = 1;
                oItem.ToPane = 1;

                // oItem.LinkTo = "EditText1"; Serve para linkar o objeto com outro.

                oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));

                oEditText.DataBind.SetBound(true, "", "EditTextF");

                // Adicionando Label
                oItem = oForm.Items.Add("lblC", BoFormItemTypes.it_STATIC);
                oItem.Left = 450;
                oItem.Width = 180;
                oItem.Top = 130;
                oItem.Height = 19;

                // comando para deixar visivel nesse folder
                oItem.FromPane = 1;
                oItem.ToPane = 1;

                oItem.LinkTo = "txtUp"; //Serve para linkar o objeto com outro.

                oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
                oStaticText.Caption = "Caminho de Armazenamento do XML:";

                // Adicionando TextBox: oEditText

                oItem = oForm.Items.Add("txtUp", BoFormItemTypes.it_EXTEDIT);
                oItem.Left = 630;
                oItem.Width = 180;
                oItem.Top = 130;
                oItem.Height = 19;
                oItem.Enabled = false;

                // comando para deixar visivel nesse folder
                oItem.FromPane = 1;
                oItem.ToPane = 1;

                oItem.LinkTo = "lblC"; //Serve para linkar o objeto com outro.

                oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));

                oEditText.DataBind.SetBound(true, "", "EditTextDS");


                // adicionando um button para selecionar o diretorio;

                oItem = oForm.Items.Add("btnUp", BoFormItemTypes.it_BUTTON);
                oItem.Left = 813;
                oItem.Width = 120;
                oItem.Top = 130;
                oItem.Height = 19;

                oItem.FromPane = 1;
                oItem.ToPane = 1;

                oButton = ((SAPbouiCOM.Button)(oItem.Specific));
                oButton.Caption = "Selecione a pasta XML";

                oItem = oForm.Items.Add("btnSearch", BoFormItemTypes.it_BUTTON);
                oItem.Left = 870;
                oItem.Width = 60;
                oItem.Top = 49;
                oItem.Height = 19;

                oItem.FromPane = 1;
                oItem.ToPane = 1;

                oButton = ((SAPbouiCOM.Button)(oItem.Specific));
                oButton.Caption = "Pesquisar";

                oItem = oForm.Items.Add("btnResetF", BoFormItemTypes.it_BUTTON);
                oItem.Left = 870;
                oItem.Width = 60;
                oItem.Top = 79;
                oItem.Height = 19;

                oItem.FromPane = 1;
                oItem.ToPane = 1;

                oButton = ((SAPbouiCOM.Button)(oItem.Specific));
                oButton.Caption = "Limpar filtro";


                // adicionando um button para selecionar o diretorio;

                oItem = oForm.Items.Add("btnMar", BoFormItemTypes.it_BUTTON);
                oItem.Left = 52;
                oItem.Width = 90;
                oItem.Top = 150;
                oItem.Height = 19;

                oItem.FromPane = 1;
                oItem.ToPane = 1;

                oButton = ((SAPbouiCOM.Button)(oItem.Specific));
                oButton.Caption = "Marcar todos";


                oItem = oForm.Items.Add("btnDes", BoFormItemTypes.it_BUTTON);
                oItem.Left = 150;
                oItem.Width = 90;
                oItem.Top = 150;
                oItem.Height = 19;

                oItem.FromPane = 1;
                oItem.ToPane = 1;

                oButton = ((SAPbouiCOM.Button)(oItem.Specific));
                oButton.Caption = "Desmarcar todos";



                ((SAPbouiCOM.Folder)(oForm.Items.Item("Conteudo").Specific)).Select(); // vai setar para acessar a primeira ABA.

                // Criando componentes da segunda aba(Configurações)

                // Adicionando Label
                oItem = oForm.Items.Add("lblSegE", BoFormItemTypes.it_STATIC);
                oItem.Left = 30;
                oItem.Width = 150;
                oItem.Top = 49;
                oItem.Height = 19;

                // comando para deixar visivel nesse folder
                oItem.FromPane = 2;
                oItem.ToPane = 2;

                oItem.LinkTo = "txtems"; //Serve para linkar o objeto com outro.

                oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
                oStaticText.Caption = "E-mail seguradora:";

                // Adicionando combobox

                oItem = oForm.Items.Add("txtems", BoFormItemTypes.it_EXTEDIT);
                oItem.Left = 130;
                oItem.Width = 300;
                oItem.Top = 49;
                oItem.Height = 19;

                // comando para deixar visivel nesse folder
                oItem.FromPane = 2;
                oItem.ToPane = 2;

                oItem.LinkTo = "lblSegE"; //Serve para linkar o objeto com outro.

                oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));

                oEditText.DataBind.SetBound(true, "", "EditTextES");


                // Adicionando Label
                oItem = oForm.Items.Add("lblEmitE", BoFormItemTypes.it_STATIC);
                oItem.Left = 30;
                oItem.Width = 150;
                oItem.Top = 89;
                oItem.Height = 19;

                // comando para deixar visivel nesse folder
                oItem.FromPane = 2;
                oItem.ToPane = 2;

                oItem.LinkTo = "txteme"; //Serve para linkar o objeto com outro.

                oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
                oStaticText.Caption = "E-mail emitente:";

                // Adicionando combobox

                oItem = oForm.Items.Add("txteme", BoFormItemTypes.it_EXTEDIT);
                oItem.Left = 130;
                oItem.Width = 300;
                oItem.Top = 89;
                oItem.Height = 19;

                // comando para deixar visivel nesse folder
                oItem.FromPane = 2;
                oItem.ToPane = 2;

                oItem.LinkTo = "lblEmitE"; //Serve para linkar o objeto com outro.

                oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));

                oEditText.DataBind.SetBound(true, "", "EditTextEE");


                // Adicionando Label
                oItem = oForm.Items.Add("lblpas", BoFormItemTypes.it_STATIC);
                oItem.Left = 30;
                oItem.Width = 100;
                oItem.Top = 129;
                oItem.Height = 19;

                // comando para deixar visivel nesse folder
                oItem.FromPane = 2;
                oItem.ToPane = 2;

                oItem.LinkTo = "txtpas"; //Serve para linkar o objeto com outro.

                oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
                oStaticText.Caption = "Senha:";

                // Adicionando combobox

                oItem = oForm.Items.Add("txtpas", BoFormItemTypes.it_EXTEDIT);
                oItem.Left = 130;
                oItem.Width = 300;
                oItem.Top = 129;
                oItem.Height = 19;

                // comando para deixar visivel nesse folder
                oItem.FromPane = 2;
                oItem.ToPane = 2;

                oItem.LinkTo = "lblpas"; //Serve para linkar o objeto com outro.

                oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));
                oEditText.IsPassword = true;

                oEditText.DataBind.SetBound(true, "", "EditPS");

                //                ---------------------------------------------------------------------

                // Adicionando Label
                oItem = oForm.Items.Add("lblhost", BoFormItemTypes.it_STATIC);
                oItem.Left = 30;
                oItem.Width = 100;
                oItem.Top = 169;
                oItem.Height = 19;

                // comando para deixar visivel nesse folder
                oItem.FromPane = 2;
                oItem.ToPane = 2;

                oItem.LinkTo = "txthost"; //Serve para linkar o objeto com outro.

                oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
                oStaticText.Caption = "Host SMTP:";

                // Adicionando combobox

                oItem = oForm.Items.Add("txthost", BoFormItemTypes.it_EXTEDIT);
                oItem.Left = 130;
                oItem.Width = 300;
                oItem.Top = 169;
                oItem.Height = 19;

                // comando para deixar visivel nesse folder
                oItem.FromPane = 2;
                oItem.ToPane = 2;

                oItem.LinkTo = "lblhost"; //Serve para linkar o objeto com outro.

                oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));

                oEditText.DataBind.SetBound(true, "", "EditHT");

                //----------------------------------------------


                // Adicionando Label
                oItem = oForm.Items.Add("lblass", BoFormItemTypes.it_STATIC);
                oItem.Left = 30;
                oItem.Width = 100;
                oItem.Top = 209;
                oItem.Height = 19;

                // comando para deixar visivel nesse folder
                oItem.FromPane = 2;
                oItem.ToPane = 2;

                oItem.LinkTo = "txtass"; //Serve para linkar o objeto com outro.

                oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
                oStaticText.Caption = "Assunto:";

                // Adicionando combobox

                oItem = oForm.Items.Add("txtass", BoFormItemTypes.it_EXTEDIT);
                oItem.Left = 130;
                oItem.Width = 300;
                oItem.Top = 209;
                oItem.Height = 19;

                // comando para deixar visivel nesse folder
                oItem.FromPane = 2;
                oItem.ToPane = 2;

                oItem.LinkTo = "lblass"; //Serve para linkar o objeto com outro.

                oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));

                oEditText.DataBind.SetBound(true, "", "EditSAS");


                // Adicionando Label
                oItem = oForm.Items.Add("lblsmtp", BoFormItemTypes.it_STATIC);
                oItem.Left = 30;
                oItem.Width = 100;
                oItem.Top = 249;
                oItem.Height = 19;

                // comando para deixar visivel nesse folder
                oItem.FromPane = 2;
                oItem.ToPane = 2;

                oItem.LinkTo = "txtsmt"; //Serve para linkar o objeto com outro.

                oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
                oStaticText.Caption = "Porta SMTP:";

                // Adicionando combobox

                oItem = oForm.Items.Add("txtsmt", BoFormItemTypes.it_EXTEDIT);
                oItem.Left = 130;
                oItem.Width = 100;
                oItem.Top = 249;
                oItem.Height = 19;

                // comando para deixar visivel nesse folder
                oItem.FromPane = 2;
                oItem.ToPane = 2;

                oItem.LinkTo = "lblsmtp"; //Serve para linkar o objeto com outro.

                oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));

                oEditText.DataBind.SetBound(true, "", "EditTextSM");


                oItem = oForm.Items.Add("optBtn", BoFormItemTypes.it_OPTION_BUTTON);
                oItem.Left = 30;
                oItem.Width = 120;
                oItem.Top = 289;
                oItem.Height = 19;
                oItem.FromPane = 2;
                oItem.ToPane = 2;
                oPtionBtn = ((SAPbouiCOM.OptionBtn)(oItem.Specific));
                oPtionBtn.Caption = "Ativar SSL";
                oPtionBtn.Selected = true;

                oPtionBtn.DataBind.SetBound(true, "", "OptDS");

                



                oItem = oForm.Items.Add("optBtD", BoFormItemTypes.it_OPTION_BUTTON);
                oItem.Left = 30;
                oItem.Width = 120;
                oItem.Top = 309;
                oItem.Height = 19;
                oItem.FromPane = 2;
                oItem.ToPane = 2;
                oPtionBtn = ((SAPbouiCOM.OptionBtn)(oItem.Specific));
                oPtionBtn.Caption = "Desativar SSL";
                oPtionBtn.GroupWith("optBtn");

                //((SAPbouiCOM.OptionBtn)(oForm.Items.Item("optBtD").Specific)).Selected = true;


                // Adicionando Label
                oItem = oForm.Items.Add("lblmes", BoFormItemTypes.it_STATIC);
                oItem.Left = 530;
                oItem.Width = 130;
                oItem.Top = 49;
                oItem.Height = 19;

                // comando para deixar visivel nesse folder
                oItem.FromPane = 2;
                oItem.ToPane = 2;

                oItem.LinkTo = "txtms"; //Serve para linkar o objeto com outro.

                oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
                oStaticText.Caption = "Mensagem corpo do e-mail:";

                // Adicionando combobox

                oItem = oForm.Items.Add("txtms", BoFormItemTypes.it_EXTEDIT);
                oItem.Left = 680;
                oItem.Width = 300;
                oItem.Top = 49;
                oItem.Height = 100;

                // comando para deixar visivel nesse folder
                oItem.FromPane = 2;
                oItem.ToPane = 2;

                oItem.LinkTo = "lblmes"; //Serve para linkar o objeto com outro.

                oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));

                oEditText.DataBind.SetBound(true, "", "EditTextMM");

                // adicionando um button para selecionar o diretorio;

                oItem = oForm.Items.Add("btnSav", BoFormItemTypes.it_BUTTON);
                oItem.Left = 30;
                oItem.Width = 60;
                oItem.Top = 349;
                oItem.Height = 19;

                oItem.FromPane = 2;
                oItem.ToPane = 2;

                oButton = ((SAPbouiCOM.Button)(oItem.Specific));
                oButton.Caption = "Salvar";

                // adicionando um button para selecionar o diretorio;

                oItem = oForm.Items.Add("btnRes", BoFormItemTypes.it_BUTTON);
                oItem.Left = 100;
                oItem.Width = 60;
                oItem.Top = 349;
                oItem.Height = 19;

                oItem.FromPane = 2;
                oItem.ToPane = 2;

                oButton = ((SAPbouiCOM.Button)(oItem.Specific));
                oButton.Caption = "Limpar";



                // adicionando um grid
                oItem = oForm.Items.Add("grid", SAPbouiCOM.BoFormItemTypes.it_GRID);
                oItem.Left = 52;
                oItem.Top = 180;
                oItem.Width = 900;
                oItem.Height = 350;

                // comando para deixar visivel nesse folder
                oItem.FromPane = 1;
                oItem.ToPane = 1;

                oForm.PaneLevel = 1;

                // Preenchendo as datas no txtDt
                DateTime today = DateTime.Now;
                today = today.AddDays(-30);

                oForm.DataSources.UserDataSources.Item("EditTextI").Value = today.ToString("yyyyMMdd");
                oForm.DataSources.UserDataSources.Item("EditTextF").Value = DateTime.Now.ToString("yyyyMMdd");

                oApplication.ItemEvent += OApplication_ItemEvent;

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message, "Erro", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
                System.Environment.Exit(0);
            }



        }

        private void preencheGrid(string filial, string dateIni, string dateFin)
        {

            oGrid = ((SAPbouiCOM.Grid)(oItem.Specific));
          

            if (oGrid.DataTable != null)
            {
                oGrid.DataTable.Clear();
            }
            else
            {
                oForm.DataSources.DataTables.Add("DataTable");
            }


            //var dateIni = oForm.DataSources.UserDataSources.Item("EditTextI").ValueEx;
            //var dateFin = oForm.DataSources.UserDataSources.Item("EditTextF").ValueEx;
            string status = "4";
            //string filial = "1";



            string query = "CALL LISTAGEMNOTAS(" + dateIni + ", " + dateFin + " , " + status + ", " + filial + " )";
            //string query = "CALL TestProc1(" + dateIni + ", " + dateFin + ")"

            oForm.DataSources.DataTables.Item(0).ExecuteQuery(query);
            oGrid.DataTable = oForm.DataSources.DataTables.Item("DataTable");

            oGrid.Columns.Item(0).Width = 60;
            oGrid.Columns.Item(1).Width = 120;
            oGrid.Columns.Item(2).Width = 100;
            oGrid.Columns.Item(3).Width = 90;
            oGrid.Columns.Item(4).Width = 100;
            oGrid.Columns.Item(5).Width = 100;
            oGrid.Columns.Item(6).Width = 50;
            oGrid.Columns.Item(7).Width = 240;



            // setando para as colunas não serem editaveis.

            oGrid.Columns.Item(1).Editable = false;
            oGrid.Columns.Item(2).Editable = false;
            oGrid.Columns.Item(3).Editable = false;
            oGrid.Columns.Item(4).Editable = false;
            oGrid.Columns.Item(5).Editable = false;
            oGrid.Columns.Item(6).Editable = false;
            oGrid.Columns.Item(7).Editable = false;

            oGrid.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;

            // Fazendo uma coluna um botão para abrir o documento.

            SAPbouiCOM.EditTextColumn oEditCol;
            oEditCol = ((SAPbouiCOM.EditTextColumn)(oGrid.Columns.Item(3)));
            oEditCol.LinkedObjectType = "2";
        }

        private void OApplication_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (FormUID.Equals("frmWebServXml"))
            {
                oForm = oApplication.Forms.Item(FormUID);
                switch (pVal.EventType)
                {
                    case BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID.Equals("Conteudo"))
                        {
                            oForm.PaneLevel = 1;
                        } else if (pVal.ItemUID.Equals("Conf"))
                        {
                            oForm.PaneLevel = 2;
                        }
                        break;
                    case BoEventTypes.et_FORM_RESIZE:
                        oForm.Freeze(true);

                        oForm.Freeze(false);
                        oForm.Update();
                        break;
                    case BoEventTypes.et_FORM_CLOSE:
                        System.Windows.Forms.Application.Exit();
                        break;
                }
            }

            if (FormUID.Equals("frmWebServXml") & pVal.BeforeAction == false & pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & pVal.ItemUID.Equals("btnResetF"))
            {
                limparFiltro();

            }
            //btnSav

            if (FormUID.Equals("frmWebServXml") & pVal.BeforeAction == false & pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & pVal.ItemUID.Equals("btnSav"))
            {
                if(oForm.DataSources.UserDataSources.Item("OptDS").Value == "")
                {
                    oApp.StatusBar.SetText("Selecione uma opção SSL", BoMessageTime.bmt_Medium,
                        BoStatusBarMessageType.smt_Error);
                    return;
                }


                saveConfig(oForm.DataSources.UserDataSources.Item("EditTextES").ValueEx, oForm.DataSources.UserDataSources.Item("EditTextEE").ValueEx,
                    oForm.DataSources.UserDataSources.Item("EditTextSM").ValueEx, oForm.DataSources.UserDataSources.Item("EditTextMM").ValueEx,
                    oForm.DataSources.UserDataSources.Item("EditSAS").ValueEx, oForm.DataSources.UserDataSources.Item("EditPS").ValueEx,
                    oForm.DataSources.UserDataSources.Item("OptDS").Value , oForm.DataSources.UserDataSources.Item("EditHT").ValueEx);


            }


            if (FormUID.Equals("frmWebServXml") & pVal.BeforeAction == false & pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & pVal.ItemUID.Equals("btnSearch"))
            {
                // validar se a data inicio não é maior que a data final.

                if (oForm.DataSources.UserDataSources.Item("ComboBoxDS").ValueEx == "0")
                {
                    oApp.StatusBar.SetText("Necessário selecionar uma filial", BoMessageTime.bmt_Medium,
                        BoStatusBarMessageType.smt_Error);
                    return;
                }

                var dataIn = DateTime.Parse(oForm.DataSources.UserDataSources.Item("EditTextI").Value);
                var dataFin = DateTime.Parse(oForm.DataSources.UserDataSources.Item("EditTextF").Value);


                if (Convert.ToDateTime(dataIn)  > Convert.ToDateTime(dataFin))
                {
                    oApp.StatusBar.SetText("Data inicial não pode ser maior que a data final", BoMessageTime.bmt_Medium,
                        BoStatusBarMessageType.smt_Error);
                    return;
                }


                preencheGrid(oForm.DataSources.UserDataSources.Item("ComboBoxDS").ValueEx, oForm.DataSources.UserDataSources.Item("EditTextI").ValueEx, oForm.DataSources.UserDataSources.Item("EditTextF").ValueEx);

            }

            if (FormUID.Equals("frmWebServXml") & pVal.BeforeAction == false & pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & pVal.ItemUID.Equals("btnRes"))
            {
                limpaConfig();

            }

            // Evento vai occorer no momento que trocar a filial no combobox
            if (FormUID.Equals("frmWebServXml") & pVal.BeforeAction == false & pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT & pVal.ItemUID.Equals("cmb"))
            {

                oForm.Items.Item("btnUp").Enabled = true;

                // verifica se o valor selecionado é diferente de zero;
                if (oForm.DataSources.UserDataSources.Item("ComboBoxDS").ValueEx != "0")
                {
                    preencheUrlFilial(oForm.DataSources.UserDataSources.Item("ComboBoxDS").ValueEx);
                }
                else
                {
                    oForm.DataSources.UserDataSources.Item("EditTextDS").ValueEx = "";
                }

            }



            if (FormUID.Equals("frmWebServXml") & pVal.BeforeAction == false & pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE)
            {

                // Preenche com base nos dados salvos de configuração.
                preencheForm();


                // Preenche combo.
                preencheCombo();




                oApp.SetStatusBarMessage(string.Format("Addon pronto para uso",
                System.Windows.Forms.Application.ProductName),
                BoMessageTime.bmt_Medium, false);

            }

            //if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST & pVal.FormType.Equals("frmWebServXml"))
            //{
            //    BeforeAction ou After
            //    if (pVal.ItemUID.Equals(""))
            //    {

            //    }

            //}

            // if (oGrid != null)

            if (FormUID.Equals("frmWebServXml") & (!pVal.BeforeAction) & (pVal.ItemUID.Equals("btnMar")) & (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED))
            {
                if (oGrid != null)
                {
                    if (oGrid.Rows.Count > 0)
                    {

                        for (int i = 0; i < oGrid.Rows.Count; i++)
                        {
                            oGrid.DataTable.SetValue("Selecione", i, "Y");
                        }
                    }
                }
                            

                            
            }

            if (FormUID.Equals("frmWebServXml") & (!pVal.BeforeAction) & (pVal.ItemUID.Equals("btnDes")) & (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED))
            {
                        if (oGrid != null)
                        {
                            if (oGrid.Rows.Count > 0)
                            {
                                for (int i = 0; i < oGrid.Rows.Count; i++)
                                {

                                    oGrid.DataTable.SetValue("Selecione", i, "");
                                }
                            }
                        }
            }

                // função quando o botão para selecionar a pasta for pressionado.
                if (FormUID.Equals("frmWebServXml") & (!pVal.BeforeAction) & (pVal.ItemUID.Equals("btnSend")) & (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED))
            {

                // verificar se foi preenchido as configurações que realizam o envio por email.

                if (oForm.DataSources.UserDataSources.Item("EditTextDS").ValueEx == "")
                {
                    oApp.StatusBar.SetText("Não foi preenchido o caminho de armazenamento", BoMessageTime.bmt_Medium,
                    BoStatusBarMessageType.smt_Error);
                    return;
                }

                if (oForm.DataSources.UserDataSources.Item("EditTextES").ValueEx == "")
                {
                    oApp.StatusBar.SetText("Não foi preenchido o e-mail da seguradora, verifique as configurações", BoMessageTime.bmt_Medium,
                    BoStatusBarMessageType.smt_Error);
                    return;
                }

                if (oForm.DataSources.UserDataSources.Item("EditTextEE").ValueEx == "")
                {
                    oApp.StatusBar.SetText("Não foi preenchido o e-mail emitente, verifique as configurações", BoMessageTime.bmt_Medium,
                    BoStatusBarMessageType.smt_Error);
                    return;
                }

                if (oForm.DataSources.UserDataSources.Item("EditTextSM").ValueEx == "")
                {
                    oApp.StatusBar.SetText("Não foi preenchido o SMTP, verifique as configurações", BoMessageTime.bmt_Medium,
                    BoStatusBarMessageType.smt_Error);
                    return;
                }

                if (oForm.DataSources.UserDataSources.Item("EditPS").ValueEx == "")
                {
                    oApp.StatusBar.SetText("Não foi preenchido o password, verifique as configurações", BoMessageTime.bmt_Medium,
                    BoStatusBarMessageType.smt_Error);
                    return;
                }

                if (oForm.DataSources.UserDataSources.Item("EditHT").ValueEx == "")
                {
                    oApp.StatusBar.SetText("Não foi preenchido o host SMTP, verifique as configurações", BoMessageTime.bmt_Medium,
                    BoStatusBarMessageType.smt_Error);
                    return;
                }

                // pegar todos os documentos selecionador e vai enviar o XML com base no endereço.

                if (oGrid != null)
                {
                    if (oGrid.Rows.Count > 0)
                    {
                        // array para armazenar a chave de acesso das notas que vão ser enviadas por email
                        List<string> listChaveAcesso = new List<string>();
                        for (int i = 0; i < oGrid.Rows.Count; i++)
                        {
                            Console.WriteLine(oGrid.DataTable.GetValue("Selecione", i));

                            if(oGrid.DataTable.GetValue("Selecione", i) == "Y")
                            {
                                string textUrl = oForm.DataSources.UserDataSources.Item("EditTextDS").ValueEx + "\\"
                                    + "procNFe" + oGrid.DataTable.GetValue("N° Chave de Acesso", i) + ".xml";
                                textUrl = textUrl.Replace("\"","");
                                listChaveAcesso.Add(textUrl);
                            }


                        }

                        // verifica se a lista de chave de acesso é maior que zero e se for será enviado o email.
                        if (listChaveAcesso.Count > 0)
                        {
                            string assunto = "";
                            if(oForm.DataSources.UserDataSources.Item("EditSAS").ValueEx != "")
                            {
                                assunto = oForm.DataSources.UserDataSources.Item("EditSAS").ValueEx;
                            }
                            else
                            {
                                assunto = "Envio de XML";
                            }

                            bool smtp = false;
                            if(oForm.DataSources.UserDataSources.Item("OptDS").ValueEx == "1")
                            {
                                smtp = true;
                            }
                            sendEmail(oForm.DataSources.UserDataSources.Item("EditTextEE").ValueEx,
                                oForm.DataSources.UserDataSources.Item("EditPS").ValueEx,
                                Convert.ToInt32(oForm.DataSources.UserDataSources.Item("EditTextSM").ValueEx),
                                oForm.DataSources.UserDataSources.Item("EditTextES").ValueEx, listChaveAcesso, assunto,
                                oForm.DataSources.UserDataSources.Item("EditTextMM").ValueEx , 
                                oForm.DataSources.UserDataSources.Item("EditHT").ValueEx
                                , smtp);

                        }
                        else
                        {
                            oApp.StatusBar.SetText("Não foi encontrado nenhum valor selecionado na grid para ser enviado", BoMessageTime.bmt_Medium,
                            BoStatusBarMessageType.smt_Error);
                            return;

                        }
                    }
                    else
                    {
                        oApp.StatusBar.SetText("Não foi encontrado valores na grid para serem enviados", BoMessageTime.bmt_Medium,
                        BoStatusBarMessageType.smt_Error);
                        return;
                    }


                }
            }



                if (FormUID.Equals("frmWebServXml") & (!pVal.BeforeAction) & (pVal.ItemUID.Equals("btnUp")) & (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED))
            {

                //SAPbouiCOM.ComboBox ComboBoxDS = (SAPbouiCOM.ComboBox)oForm.Items.Item("cmb").Specific;

                if(oForm.DataSources.UserDataSources.Item("ComboBoxDS").ValueEx == "0")
                {
                    System.Windows.Forms.MessageBox.Show("É obrigatório selecionar uma filial" , "Erro", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
                    return;
                }

                oForm.Items.Item("btnUp").Enabled = false;
               

                Thread t = new Thread(() =>
                {
                    var form = new System.Windows.Forms.Form();

                    FolderBrowserDialog dialog = new FolderBrowserDialog();

                    dialog.Description = "Por favor seleciona uma pasta";
                    dialog.RootFolder = Environment.SpecialFolder.MyComputer;

                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        // = dialog.SelectedPath;
                        //txtUp.value = dialog.SelectedPath;
                        oForm.DataSources.UserDataSources.Item("EditTextDS").ValueEx = dialog.SelectedPath;
                        saveConfigFilial(oForm.DataSources.UserDataSources.Item("ComboBoxDS").ValueEx, oForm.DataSources.UserDataSources.Item("EditTextDS").ValueEx);
                    }
                    else
                    {
                        oForm.DataSources.UserDataSources.Item("EditTextDS").ValueEx = "";
                        saveConfigFilial(oForm.DataSources.UserDataSources.Item("ComboBoxDS").ValueEx, oForm.DataSources.UserDataSources.Item("EditTextDS").ValueEx);
                    }

                    form.Close();
                });         

                t.SetApartmentState(ApartmentState.STA);
                t.Start();

            }

        }

        private void Open()
        {
            Thread t = new Thread(new ThreadStart(this.DialogoSelecaoArquivo));
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
        }


        private void limparFiltro()
        {
            SAPbouiCOM.ComboBox ComboBoxDS = (SAPbouiCOM.ComboBox)oForm.Items.Item("cmb").Specific;
            ComboBoxDS.Select("0", SAPbouiCOM.BoSearchKey.psk_ByValue);

            oForm.Items.Item("btnUp").Enabled = true;

            // Preenchendo as datas no txtDt
            DateTime today = DateTime.Now;
            today = today.AddDays(-30);

            oForm.DataSources.UserDataSources.Item("EditTextI").Value = today.ToString("yyyyMMdd");
            oForm.DataSources.UserDataSources.Item("EditTextF").Value = DateTime.Now.ToString("yyyyMMdd");
            oForm.DataSources.UserDataSources.Item("EditTextDS").Value = ""; 
        }

        private void sendEmail(string emailEmitente , string pass, int portaSmtp, string emailDest, List<string> caminhoFile, string assunto , string mensagem , string hostSmtp , bool ssl)
        {
            try
            {
                oApp.SetStatusBarMessage(string.Format("Addon Alfa Seguradora - Enviando email para: " + emailDest,
                System.Windows.Forms.Application.ProductName),
                BoMessageTime.bmt_Medium, false);


                using (SmtpClient smtp = new SmtpClient())
                {
                    using (MailMessage email = new MailMessage())
                    {
                        // Servidor SMTP
                        smtp.Host = hostSmtp;
                        smtp.UseDefaultCredentials = false;
                        smtp.Credentials = new System.Net.NetworkCredential(emailEmitente, pass);
                        smtp.Port = portaSmtp;
                        smtp.EnableSsl = ssl;
                        
                        email.From = new MailAddress(emailEmitente);
                        email.To.Add(emailDest);

                        email.Subject = assunto;
                        email.IsBodyHtml = false;
                        email.Body = mensagem;

                        foreach(string url in caminhoFile)
                        {
                            email.Attachments.Add(new System.Net.Mail.Attachment(url));
                        }
                        

                        smtp.Send(email);

                        oApp.SetStatusBarMessage(string.Format("Addon Alfa Seguradora - Concluído envio de e-mail. ",
                        System.Windows.Forms.Application.ProductName),
                        BoMessageTime.bmt_Medium, false);
                    }
                }
            }catch(Exception ex)
            {
                oApp.SetStatusBarMessage(string.Format("Addon Alfa Seguradora - Houve um erro ao enviar email para: " + emailDest + 
                    " Motivo do erro:" + ex.Message,
                System.Windows.Forms.Application.ProductName),
                BoMessageTime.bmt_Medium, false);
            }
        }

        private void preencheForm()
        {
            // realizar consulta na tabela de configuração e trazer aqui os valores cadastrados lá.

            oApp.SetStatusBarMessage(string.Format("Addon - Buscando configurações...",
                    System.Windows.Forms.Application.ProductName),
                    BoMessageTime.bmt_Medium, false);


            oRecordset = (SAPbobsCOM.Recordset)_company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = "SELECT \"U_email_seguradora\", \"U_email_emitente\", \"U_smtp\", \"U_mensagem\" ,\"U_assunto\" " +
                ", \"U_pass\" , \"U_atvSSL\" , \"U_host\" FROM  \"@SEGURADORA_CONF\" " +
                "WHERE \"Code\" = 'ADDON_ALFA_SEGURADORA' AND \"Name\" = 'ADDON_ALFA_SEGURADORA'";


            oRecordset.DoQuery(query);
            if (oRecordset.RecordCount > 0)
            {
                while (!oRecordset.EoF)
                {
                    oForm.DataSources.UserDataSources.Item("EditTextES").ValueEx = oRecordset.Fields.Item("U_email_seguradora").Value.ToString();
                    oForm.DataSources.UserDataSources.Item("EditTextEE").ValueEx = oRecordset.Fields.Item("U_email_emitente").Value.ToString();
                    oForm.DataSources.UserDataSources.Item("EditTextSM").ValueEx = oRecordset.Fields.Item("U_smtp").Value.ToString();
                    oForm.DataSources.UserDataSources.Item("EditTextMM").ValueEx = oRecordset.Fields.Item("U_mensagem").Value.ToString();
                    oForm.DataSources.UserDataSources.Item("EditSAS").ValueEx = oRecordset.Fields.Item("U_assunto").Value.ToString();
                    oForm.DataSources.UserDataSources.Item("EditPS").ValueEx = oRecordset.Fields.Item("U_pass").Value.ToString();
                    oForm.DataSources.UserDataSources.Item("EditHT").ValueEx = oRecordset.Fields.Item("U_host").Value.ToString();

                    // option 
                    if (oRecordset.Fields.Item("U_atvSSL").Value.ToString() == "1")
                    {
                        ((SAPbouiCOM.OptionBtn)(oForm.Items.Item("optBtn").Specific)).Selected = true;
                        oForm.DataSources.UserDataSources.Item("OptDS").ValueEx = "1";
                        //oForm.DataSources.UserDataSources.Item("optBtn").ValueEx = ;
                    }
                    else
                    {
                        ((SAPbouiCOM.OptionBtn)(oForm.Items.Item("optBtD").Specific)).Selected = true;
                        oForm.DataSources.UserDataSources.Item("OptDS").ValueEx = "2";
                    }
                    
                    break;
                }
            }
            else
            {
                limpaConfig();
            }
        }



        private void preencheUrlFilial(string filial)
        {
            string query = "SELECT \"U_filial\", \"U_diretorio\" FROM \"@SEG_CONF_FILIAIS\" " +
                "WHERE \"Code\" = 'ADDON_ALFA_SEGURADORA' AND \"Name\" = 'ADDON_ALFA_SEGURADORA' AND " +
                    "  \"U_filial\" = '" + filial + "'";


            oRecordset.DoQuery(query);
            if (oRecordset.RecordCount > 0)
            {
                while (!oRecordset.EoF)
                {
                    oForm.DataSources.UserDataSources.Item("EditTextDS").ValueEx = oRecordset.Fields.Item("U_diretorio").Value.ToString();
                    break;
                }
            }
            else
            {
                oForm.DataSources.UserDataSources.Item("EditTextDS").ValueEx = "";
            }
        }


        private void saveConfigFilial(string filial , string url)
        {

            string query = "SELECT \"U_filial\", \"U_diretorio\" FROM \"@SEG_CONF_FILIAIS\" " +
                "WHERE \"Code\" = 'ADDON_ALFA_SEGURADORA' AND \"Name\" = 'ADDON_ALFA_SEGURADORA' AND " +
                "  \"U_filial\" = '" + filial + "'";


            oRecordset.DoQuery(query);
            if (oRecordset.RecordCount > 0)
            {
                string update = "UPDATE \"@SEG_CONF_FILIAIS\" SET \"U_diretorio\" = '"+ url+ "' " +
                    "WHERE \"Code\" = 'ADDON_ALFA_SEGURADORA' AND \"Name\" = 'ADDON_ALFA_SEGURADORA' AND " +
                "  \"U_filial\" = '" + filial + "'";

                oRecordset.DoQuery(update);
            }
            else
            {
                string insert = "INSERT INTO \"@SEG_CONF_FILIAIS\" (\"U_filial\", \"U_diretorio\" ," +
                    " \"Code\", \"Name\") VALUES ('" + filial + "' , '"+ url +"' , 'ADDON_ALFA_SEGURADORA' , 'ADDON_ALFA_SEGURADORA')";
                oRecordset.DoQuery(insert);
            }

        }

        private void limpaConfig()
        {
            oForm.DataSources.UserDataSources.Item("EditTextES").ValueEx = "";
            oForm.DataSources.UserDataSources.Item("EditTextEE").ValueEx = "";
            oForm.DataSources.UserDataSources.Item("EditTextSM").ValueEx = "";
            oForm.DataSources.UserDataSources.Item("EditTextMM").ValueEx = "";
            oForm.DataSources.UserDataSources.Item("EditSAS").ValueEx = "";
            oForm.DataSources.UserDataSources.Item("EditPS").ValueEx = "";
            oForm.DataSources.UserDataSources.Item("EditHT").ValueEx = "";
            ((SAPbouiCOM.OptionBtn)(oForm.Items.Item("optBtn").Specific)).Selected = true;
        }

        private void saveConfig(string email_seguradora , string email_emitente, string smtp, string mensagem , string assunto, string pass, string ssl , string hostSmtp)
        {


            oApp.SetStatusBarMessage(string.Format("Addon - Salvando configurações...",
            System.Windows.Forms.Application.ProductName),
            BoMessageTime.bmt_Medium, false);


            oRecordset = (SAPbobsCOM.Recordset)_company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = "SELECT \"U_email_seguradora\", \"U_email_emitente\", \"U_smtp\", \"U_mensagem\" ,\"U_assunto\" " +
                ", \"U_pass\" , \"U_atvSSL\" , \"U_host\"  FROM  \"@SEGURADORA_CONF\" " +
                "WHERE \"Code\" = 'ADDON_ALFA_SEGURADORA' AND \"Name\" = 'ADDON_ALFA_SEGURADORA'";


            oRecordset.DoQuery(query);
            if (oRecordset.RecordCount > 0)
            {
                // Executa Update;

                string update = "UPDATE \"@SEGURADORA_CONF\" SET \"U_email_seguradora\" = '" + email_seguradora + "'," +
                                " \"U_email_emitente\" = '" + email_emitente + "'," +
                                " \"U_smtp\" = '" + smtp + "'," +
                                " \"U_mensagem\" = '" + mensagem + "'," +
                                " \"U_assunto\" = '" + assunto + "'," +
                                " \"U_pass\" = '" + pass + "'," +
                                " \"U_atvSSL\" = '" + ssl + "'," +
                                " \"U_host\" = '" + hostSmtp + "'" +
                                " WHERE \"Code\" = 'ADDON_ALFA_SEGURADORA' AND \"Name\" = 'ADDON_ALFA_SEGURADORA'";

                oRecordset.DoQuery(update);
            }
            else
            {
                // Executa Insert;

                string insert = "INSERT INTO \"@SEGURADORA_CONF\" " +
                            "(\"U_email_seguradora\", \"U_email_emitente\", \"U_smtp\", \"U_mensagem\", \"U_assunto\", \"U_pass\"," +
                            " \"U_atvSSL\", \"U_host\",  \"Code\", \"Name\") " +
                            "VALUES('" + email_seguradora + "', '" + email_emitente + "', '" + smtp + "'," +
                            " '" + mensagem + "', '" + assunto + "', '" + pass + "', '" + ssl + "','" + hostSmtp + "'," +
                            " 'ADDON_ALFA_SEGURADORA', 'ADDON_ALFA_SEGURADORA')";

                oRecordset.DoQuery(insert);
            }


            oApp.SetStatusBarMessage(string.Format("Addon - Configuração salva com sucesso.",
            System.Windows.Forms.Application.ProductName),
            BoMessageTime.bmt_Medium, false);


        }

        private void preencheCombo()
        {
            oRecordset = (SAPbobsCOM.Recordset)_company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = "SELECT \"BPLId\" ,\"BPLName\"  FROM OBPL WHERE \"Disabled\" = 'N'";

            oRecordset.DoQuery(query);
            if (oRecordset.RecordCount > 0)
            {
                while (!oRecordset.EoF)
                {
                    for(int i = 0; i < oRecordset.RecordCount; i++)
                    {
                        SAPbouiCOM.ComboBox ComboBoxDS = (SAPbouiCOM.ComboBox)oForm.Items.Item("cmb").Specific;

                        if(i == 0)
                        {
                            ComboBoxDS.ValidValues.Add("0", "Selecione uma filial");
                            ComboBoxDS.Select("0", SAPbouiCOM.BoSearchKey.psk_ByValue);
                        }
                        ComboBoxDS.ValidValues.Add(oRecordset.Fields.Item(0).Value.ToString(), oRecordset.Fields.Item(1).Value.ToString());


                    }
                    oRecordset.MoveNext();
                }
            }
        }


        private void DialogoSelecaoArquivo()
        {
            var form = new System.Windows.Forms.Form();

            FolderBrowserDialog dialog = new FolderBrowserDialog();

            dialog.Description = "Por favor seleciona uma pasta";
            dialog.RootFolder = Environment.SpecialFolder.MyComputer;

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                // = dialog.SelectedPath;
                //txtUp.value = dialog.SelectedPath;
                oForm.DataSources.UserDataSources.Item("EditTextDS").ValueEx = dialog.SelectedPath;
            }
            else
            {
                oForm.DataSources.UserDataSources.Item("EditTextDS").ValueEx = "";
            }

            form.Close();

        }

        private void AddDataSourceNoForm()
        {
            oForm.DataSources.UserDataSources.Add("FolderDS", BoDataType.dt_SHORT_TEXT, 1);
            oForm.DataSources.UserDataSources.Add("ComboBoxDS", BoDataType.dt_SHORT_TEXT, 100);
            oForm.DataSources.UserDataSources.Add("EditTextDS", BoDataType.dt_SHORT_TEXT, 100);
            oForm.DataSources.UserDataSources.Add("EditTextI", BoDataType.dt_DATE, 100);
            oForm.DataSources.UserDataSources.Add("EditTextF", BoDataType.dt_DATE, 100);
            oForm.DataSources.UserDataSources.Add("EditTextES", BoDataType.dt_SHORT_TEXT, 100);
            oForm.DataSources.UserDataSources.Add("EditTextEE", BoDataType.dt_SHORT_TEXT, 100);
            oForm.DataSources.UserDataSources.Add("EditTextSM", BoDataType.dt_SHORT_TEXT, 10);
            oForm.DataSources.UserDataSources.Add("EditTextMM", BoDataType.dt_SHORT_TEXT, 100);
            oForm.DataSources.UserDataSources.Add("EditSAS", BoDataType.dt_SHORT_TEXT, 100);
            oForm.DataSources.UserDataSources.Add("EditPS", BoDataType.dt_SHORT_TEXT, 100);
            oForm.DataSources.UserDataSources.Add("EditHT", BoDataType.dt_SHORT_TEXT, 100);
             oForm.DataSources.UserDataSources.Add("OptDS", BoDataType.dt_SHORT_TEXT, 1);
        }

        public FormWebServiceXml()
        {
            SetApplication();

            // Realiza a criação das tabelas utilizadas no addon.
            DI.Inicialize();

            createForm();

            oForm.Visible = true;
        }


    }
}
