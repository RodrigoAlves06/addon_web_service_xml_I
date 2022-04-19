using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM;

namespace AddonWebServiceXml
{
    public class FormWebServiceXml
    {
        private SAPbouiCOM.Application oApplication;
        private SAPbouiCOM.Form oForm;
        private SAPbouiCOM.Item oItem;
        private void SetApplication()
        {
            SboGuiApi oSboGuiApi = null;
            string sConnectionString = null;

            oSboGuiApi = new SAPbouiCOM.SboGuiApi();

            sConnectionString = Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));

            try
            {
                oSboGuiApi.Connect(sConnectionString);
            }
            catch(Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Erro de conexão Addon WebService XML: " + ex.Message);
                System.Environment.Exit(0);
            }

            oApplication = oSboGuiApi.GetApplication(-1);
            oApplication.SetStatusBarMessage(string.Format("Addon WebService XML Conectado com sucesso!",
                System.Windows.Forms.Application.ProductName),
                BoMessageTime.bmt_Medium, false);
        }

        private void createForm()
        {
            SAPbouiCOM.Button oButton = null;
            SAPbouiCOM.Folder oFolder = null;
            SAPbouiCOM.ComboBox oComboBox = null;

            SAPbouiCOM.FormCreationParams oCreationParams = null;
            oCreationParams = ((SAPbouiCOM.FormCreationParams)(oApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));

            oCreationParams.UniqueID = "frmWebServXml";
            oCreationParams.FormType = "frmWebServXml";
            oCreationParams.BorderStyle = BoFormBorderStyle.fbs_Sizable;

            oForm = oApplication.Forms.AddEx(oCreationParams);

            AddDataSourceNoForm();
            oForm.Title = "Addon - Envio XML seguradora";
            oForm.Left = 300;
            //oForm.ClientWidth = 200;
            //oForm.Top = 100;
            //oForm.ClientHeight = 140;

            oForm.ClientWidth = 1000;
            oForm.Top = 100;
            oForm.ClientHeight = 640;

            oItem = oForm.Items.Add("1", BoFormItemTypes.it_BUTTON);
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

            oItem = oForm.Items.Add("React", BoFormItemTypes.it_RECTANGLE);
            oItem.Left = 0;
            oItem.Width = 994;
            oItem.Top = 22;
            oItem.Height = 580;

            oItem = oForm.Items.Add("Conteudo", BoFormItemTypes.it_FOLDER);
            oItem.Left = 0;
            oItem.Width = 120;
            oItem.Top = 6;
            oItem.Height = 19;

            oFolder = ((SAPbouiCOM.Folder)(oItem.Specific));
            oFolder.Caption = "Conteúdo Envio XML";
            oFolder.DataBind.SetBound(true, "", "FolderDS");
            oFolder.Select();

            oItem = oForm.Items.Add("Conf", BoFormItemTypes.it_FOLDER);
            oItem.Left = 100;
            oItem.Width = 120;
            oItem.Top = 6;
            oItem.Height = 19;

            oFolder = ((SAPbouiCOM.Folder)(oItem.Specific));
            oFolder.Caption = "Configurações";
            oFolder.DataBind.SetBound(true, "", "FolderDS");
            oFolder.GroupWith("Conteudo");

            // Add 

        }

        private void AddDataSourceNoForm()
        {
            oForm.DataSources.UserDataSources.Add("FolderDS", BoDataType.dt_SHORT_TEXT, 1);
            oForm.DataSources.UserDataSources.Add("OptBtnDS", BoDataType.dt_SHORT_TEXT, 1);
        }

        public FormWebServiceXml()
        {
            SetApplication();

            createForm();

            oForm.Visible = true;
        }
    }
}
