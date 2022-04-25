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
                oItem.Enabled = false;
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
                oItem.Enabled = false;

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


                // adicionando um grid
                oItem = oForm.Items.Add("grid", SAPbouiCOM.BoFormItemTypes.it_GRID);
                oItem.Left = 52;
                oItem.Top = 180;
                oItem.Width = 900;
                oItem.Height = 350;

                // comando para deixar visivel nesse folder
                oItem.FromPane = 1;
                oItem.ToPane = 1;

                oGrid = ((SAPbouiCOM.Grid)(oItem.Specific));

                oForm.DataSources.DataTables.Add("DataTable");
                string query = @"SELECT 
                    '' AS ""Selecione"" 
                    , ""CardCode"", /* AS ""Data do Faturamento"",*/
                    ""CardName"" AS ""Número da NF"" 
                    , ""DocDate"" AS ""Código do PN""
                    , ""DocDate"" AS ""Nome do PN""
                    , ""DocNum"" AS ""Placa Segurado""
                    , ""DocStatus"" AS ""Status""
                    , ""DocStatus"" AS ""N° Chave de acesso""
                    FROM OINV";

                
                oForm.DataSources.DataTables.Item(0).ExecuteQuery(query);
                oGrid.DataTable = oForm.DataSources.DataTables.Item("DataTable");

                oGrid.Columns.Item(0).Width = 60;
                oGrid.Columns.Item(1).Width = 130;
                oGrid.Columns.Item(2).Width = 100;
                oGrid.Columns.Item(3).Width = 90;
                oGrid.Columns.Item(4).Width = 100;
                oGrid.Columns.Item(5).Width = 100;
                oGrid.Columns.Item(6).Width = 100;
                oGrid.Columns.Item(7).Width = 180;

                oForm.PaneLevel = 1;

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
                oEditCol = ((SAPbouiCOM.EditTextColumn)(oGrid.Columns.Item("CardCode")));
                oEditCol.LinkedObjectType = "2";


                

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



                // Preenchendo as datas no txtDt
                DateTime today = DateTime.Now;
                today = today.AddDays(-30);

                oForm.DataSources.UserDataSources.Item("EditTextI").ValueEx = today.ToString("yyyyMMdd");
                oForm.DataSources.UserDataSources.Item("EditTextF").ValueEx = DateTime.Now.ToString("yyyyMMdd");
                

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
                oItem = oForm.Items.Add("lblsmtp", BoFormItemTypes.it_STATIC);
                oItem.Left = 30;
                oItem.Width = 100;
                oItem.Top = 129;
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
                oItem.Top = 129;
                oItem.Height = 19;

                // comando para deixar visivel nesse folder
                oItem.FromPane = 2;
                oItem.ToPane = 2;

                oItem.LinkTo = "lblsmtp"; //Serve para linkar o objeto com outro.

                oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));

                oEditText.DataBind.SetBound(true, "", "EditTextSM");


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

                //---------------------------------------------------------------------------------


                // Adicionando Label
                oItem = oForm.Items.Add("lblass", BoFormItemTypes.it_STATIC);
                oItem.Left = 30;
                oItem.Width = 100;
                oItem.Top = 169;
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
                oItem.Width = 100;
                oItem.Top = 169;
                oItem.Height = 19;

                // comando para deixar visivel nesse folder
                oItem.FromPane = 2;
                oItem.ToPane = 2;

                oItem.LinkTo = "lblass"; //Serve para linkar o objeto com outro.

                oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));

                oEditText.DataBind.SetBound(true, "", "EditTextSAS");




                // Adicionando Label
                oItem = oForm.Items.Add("lblpas", BoFormItemTypes.it_STATIC);
                oItem.Left = 30;
                oItem.Width = 100;
                oItem.Top = 209;
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
                oItem.Width = 100;
                oItem.Top = 209;
                oItem.Height = 19;

                // comando para deixar visivel nesse folder
                oItem.FromPane = 2;
                oItem.ToPane = 2;

                oItem.LinkTo = "lblpas"; //Serve para linkar o objeto com outro.

                oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));

                oEditText.DataBind.SetBound(true, "", "EditTextPS");






                // adicionando um button para selecionar o diretorio;

                oItem = oForm.Items.Add("btnSav", BoFormItemTypes.it_BUTTON);
                oItem.Left = 30;
                oItem.Width = 60;
                oItem.Top = 189;
                oItem.Height = 19;

                oItem.FromPane = 2;
                oItem.ToPane = 2;

                oButton = ((SAPbouiCOM.Button)(oItem.Specific));
                oButton.Caption = "Salvar";

                // adicionando um button para selecionar o diretorio;

                oItem = oForm.Items.Add("btnRes", BoFormItemTypes.it_BUTTON);
                oItem.Left = 100;
                oItem.Width = 60;
                oItem.Top = 189;
                oItem.Height = 19;

                oItem.FromPane = 2;
                oItem.ToPane = 2;

                oButton = ((SAPbouiCOM.Button)(oItem.Specific));
                oButton.Caption = "Limpar";

                oApplication.ItemEvent += OApplication_ItemEvent;

                // Add 
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message, "Erro", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
                System.Environment.Exit(0);
            }



        }

        private void OApplication_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (FormUID.Equals("frmWebServXml"))
            {
                oForm = oApplication.Forms.Item(FormUID);
                switch(pVal.EventType)
                {
                    case BoEventTypes.et_ITEM_PRESSED:
                        if(pVal.ItemUID.Equals("Conteudo"))
                        {
                            oForm.PaneLevel = 1;
                        }else if (pVal.ItemUID.Equals("Conf"))
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

            if(FormUID.Equals("frmWebServXml") & pVal.BeforeAction == false & pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & pVal.ItemUID.Equals("btnResetF"))
            {
                limparFiltro();

            }



            if (FormUID.Equals("frmWebServXml") & pVal.BeforeAction == false  & pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE)
            {
                // Realiza a criação das tabelas utilizadas no addon.
                DI.Inicialize();

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

            // função quando o botão para selecionar a pasta for pressionado.

            if(FormUID.Equals("frmWebServXml") & (!pVal.BeforeAction) & (pVal.ItemUID.Equals("btnUp")) & (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED))
            {

                //dynamic retornoLogout = ServiceLayer.Login();

                oForm.Items.Item("btnUp").Enabled = false;
                //Open();
                // fazer uma validação antes, para verificar se foi selecionado uma filial;

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
                    }
                    else
                    {
                        oForm.DataSources.UserDataSources.Item("EditTextDS").ValueEx = "";
                    }

                    form.Close();
                });         

                t.SetApartmentState(ApartmentState.STA);
                t.Start();





                oForm.Items.Item("btnUp").Enabled = true;
                //     "C:\\";





                //Thread t = new Thread(() =>
                //{
                //    FolderBrowserDialog dialog = new FolderBrowserDialog();

                //    dialog.Description = "Por favor seleciona uma pasta";
                //    dialog.RootFolder = Environment.SpecialFolder.MyComputer;

                //    if (dialog.ShowDialog() == DialogResult.OK)
                //    {
                //        // = dialog.SelectedPath;
                //        //txtUp.value = dialog.SelectedPath;
                //        oForm.DataSources.UserDataSources.Item("EditTextDS").ValueEx = dialog.SelectedPath; 
                //    }
                //    else
                //    {
                //        oForm.DataSources.UserDataSources.Item("EditTextDS").ValueEx = "";
                //    }

                //});



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

            // Preenchendo as datas no txtDt
            DateTime today = DateTime.Now;
            today = today.AddDays(-30);

            oForm.DataSources.UserDataSources.Item("EditTextI").ValueEx = today.ToString("yyyyMMdd");
            oForm.DataSources.UserDataSources.Item("EditTextF").ValueEx = DateTime.Now.ToString("yyyyMMdd");
            oForm.DataSources.UserDataSources.Item("EditTextDS").ValueEx = ""; 
        }

        private void sendEmail(string emailEmitente , string pass, int portaSmtp, string emailDest, string caminhoFile, string assunto , string mensagem , bool ssl)
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
                        smtp.Host = emailEmitente;
                        smtp.UseDefaultCredentials = false;
                        smtp.Credentials = new System.Net.NetworkCredential(emailEmitente, pass);
                        smtp.Port = portaSmtp;
                        smtp.EnableSsl = ssl;
                        
                        email.From = new MailAddress(emailEmitente);
                        email.To.Add(emailDest);

                        email.Subject = assunto;
                        email.IsBodyHtml = false;
                        email.Body = mensagem;


                        email.Attachments.Add(new System.Net.Mail.Attachment(caminhoFile));
                        // ver a necessidade de enviar um email com todos os anexos.

                        smtp.Send(email);
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

            string query = "SELECT \"U_email_seguradora\", \"U_email_emitente\", \"U_smtp\", \"U_mensagem\" FROM " +
                "\"@SEGURADORA_CONF\"";

            oRecordset.DoQuery(query);
            if (oRecordset.RecordCount > 0)
            {
                while (!oRecordset.EoF)
                {
                    oForm.DataSources.UserDataSources.Item("EditTextES").ValueEx = oRecordset.Fields.Item("U_email_seguradora").Value.ToString();
                    oForm.DataSources.UserDataSources.Item("EditTextEE").ValueEx = oRecordset.Fields.Item("U_email_emitente").Value.ToString();
                    oForm.DataSources.UserDataSources.Item("EditTextSM").ValueEx = oRecordset.Fields.Item("smtp").Value.ToString();
                    oForm.DataSources.UserDataSources.Item("EditTextMM").ValueEx = oRecordset.Fields.Item("mensagem").Value.ToString();
                    break;
                }
            }
            else
            {
                oForm.DataSources.UserDataSources.Item("EditTextES").ValueEx = "";
                oForm.DataSources.UserDataSources.Item("EditTextEE").ValueEx = "";
                oForm.DataSources.UserDataSources.Item("EditTextSM").ValueEx = "";
                oForm.DataSources.UserDataSources.Item("EditTextMM").ValueEx = "";
            }
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
                            ComboBoxDS.ValidValues.Add("0", "Seleione uma filial");
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
            oForm.DataSources.UserDataSources.Add("EditTextSAS", BoDataType.dt_SHORT_TEXT, 100);
            oForm.DataSources.UserDataSources.Add("EditTextPS", BoDataType.dt_SHORT_TEXT, 100);
            //oForm.DataSources.UserDataSources.Add("StaticTextDS", BoDataType.dt_SHORT_TEXT, 1);
            // oForm.DataSources.UserDataSources.Add("OptBtnDS", BoDataType.dt_SHORT_TEXT, 1);
        }

        public FormWebServiceXml()
        {
            SetApplication();

            createForm();

            oForm.Visible = true;
        }


    }
}
