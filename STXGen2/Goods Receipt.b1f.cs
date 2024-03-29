
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using SAPbouiCOM;

namespace STXGen2
{

    [FormAttribute("721", "Goods Receipt.b1f")]
    class Goods_Receipt : SystemFormBase
    {
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.EditText EditText3;
        private SAPbouiCOM.EditText EditText4;
        private SAPbouiCOM.EditText EditText5;
        private SAPbouiCOM.EditText EditText6;
        private SAPbouiCOM.EditText EditText7;
        private SAPbouiCOM.EditText EditText8;
        private SAPbouiCOM.EditText EditText9;
        private SAPbouiCOM.PictureBox PictureBox0;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.EditText EditText10;
        private SAPbouiCOM.ComboBox ComboBox1;
        private SAPbouiCOM.EditText EditText11;
        private SAPbouiCOM.EditText EditText12;
        private SAPbouiCOM.EditText EditText13;
        private SAPbouiCOM.EditText EditText14;
        private SAPbouiCOM.EditText EditText15;
        private SAPbouiCOM.EditText EditText16;
        private SAPbouiCOM.EditText EditText17;
        private SAPbouiCOM.EditText EditText18;
        private SAPbouiCOM.EditText EditText19;
        private SAPbouiCOM.EditText EditText20;
        private SAPbouiCOM.EditText EditText21;
        private SAPbouiCOM.EditText EditText22;
        private SAPbouiCOM.EditText EditText23;
        private SAPbouiCOM.CheckBox CheckBox0;
        private SAPbouiCOM.CheckBox CheckBox1;
        private SAPbouiCOM.CheckBox CheckBox2;
        private SAPbouiCOM.CheckBox CheckBox3;
        private SAPbouiCOM.CheckBox CheckBox4;
        private SAPbouiCOM.CheckBox CheckBox5;
        private SAPbouiCOM.CheckBox CheckBox6;
        private SAPbouiCOM.CheckBox CheckBox7;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.StaticText StaticText3;
        private DataTable selectedDataTable;

        public Goods_Receipt()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("WONum").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("Item_2").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("Item_3").Specific));
            this.EditText4 = ((SAPbouiCOM.EditText)(this.GetItem("Item_4").Specific));
            this.EditText5 = ((SAPbouiCOM.EditText)(this.GetItem("Item_5").Specific));
            this.EditText6 = ((SAPbouiCOM.EditText)(this.GetItem("Item_6").Specific));
            this.EditText7 = ((SAPbouiCOM.EditText)(this.GetItem("Item_7").Specific));
            this.EditText8 = ((SAPbouiCOM.EditText)(this.GetItem("Item_8").Specific));
            this.EditText9 = ((SAPbouiCOM.EditText)(this.GetItem("Item_9").Specific));
            this.PictureBox0 = ((SAPbouiCOM.PictureBox)(this.GetItem("Item_10").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_11").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("btClear").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_13").Specific));
            this.EditText10 = ((SAPbouiCOM.EditText)(this.GetItem("Item_14").Specific));
            this.ComboBox1 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_15").Specific));
            this.EditText11 = ((SAPbouiCOM.EditText)(this.GetItem("Item_16").Specific));
            this.EditText12 = ((SAPbouiCOM.EditText)(this.GetItem("ToolNum").Specific));
            this.EditText13 = ((SAPbouiCOM.EditText)(this.GetItem("ToolWgt").Specific));
            this.EditText14 = ((SAPbouiCOM.EditText)(this.GetItem("ToolLen").Specific));
            this.EditText15 = ((SAPbouiCOM.EditText)(this.GetItem("Item_20").Specific));
            this.EditText16 = ((SAPbouiCOM.EditText)(this.GetItem("ToolHgt").Specific));
            this.EditText17 = ((SAPbouiCOM.EditText)(this.GetItem("Item_24").Specific));
            this.EditText18 = ((SAPbouiCOM.EditText)(this.GetItem("Item_25").Specific));
            this.EditText19 = ((SAPbouiCOM.EditText)(this.GetItem("Item_26").Specific));
            this.EditText20 = ((SAPbouiCOM.EditText)(this.GetItem("Item_27").Specific));
            this.EditText21 = ((SAPbouiCOM.EditText)(this.GetItem("Item_28").Specific));
            this.EditText22 = ((SAPbouiCOM.EditText)(this.GetItem("Item_29").Specific));
            this.EditText23 = ((SAPbouiCOM.EditText)(this.GetItem("Item_30").Specific));
            this.CheckBox0 = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_31").Specific));
            this.CheckBox1 = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_32").Specific));
            this.CheckBox2 = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_33").Specific));
            this.CheckBox3 = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_34").Specific));
            this.CheckBox4 = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_35").Specific));
            this.CheckBox5 = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_36").Specific));
            this.CheckBox6 = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_37").Specific));
            this.CheckBox7 = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_38").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_39").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_40").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_41").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_42").Specific));
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("btTFind").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.ResizeAfter += new SAPbouiCOM.Framework.FormBase.ResizeAfterHandler(this.Form_ResizeAfter);

        }

        private void OnCustomInitialize()
        {
            this.Button2 = this.GetItem("btTFind").Specific as SAPbouiCOM.Button;
            if (this.Button2 != null)
            {
                this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("btTFind").Specific));
                this.Button2.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button5_PressedAfter);
               // DisableTrfsBtn();
            }
        }

        private void Button5_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {

            SAPbouiCOM.DBDataSource dbDataSource = (SAPbouiCOM.DBDataSource)this.UIAPIRawForm.DataSources.DBDataSources.Item(0);
            Utils.ParentFormUID = this.UIAPIRawForm.UniqueID;
            string docEntry = dbDataSource.GetValue("DocEntry", 0).Trim();

            if (!IsFormOpen("ToolFindUI"))
            {
                
                formToolInfo = new ToolFind();
                formToolInfo.UIAPIRawForm.Visible = true;
            }
            else
            {
                SAPbouiCOM.Form existingForm = Program.SBO_Application.Forms.Item("ToolFindUI");
                existingForm.Visible = true;
            }
        }

        private bool IsFormOpen(string formUID)
        {
            for (int i = 0; i < Program.SBO_Application.Forms.Count; i++)
            {
                if (Program.SBO_Application.Forms.Item(i).UniqueID == formUID)
                {
                    return true;
                }
            }
            return false;
        }

        private void Form_ResizeAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            this.UIAPIRawForm.Freeze(true);
            SAPbouiCOM.EditText edtWONum = (SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("WONum").Specific;
            SAPbouiCOM.EditText edtRemarks = (SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("11").Specific;
            SAPbouiCOM.Matrix mtxContents = (SAPbouiCOM.Matrix)this.UIAPIRawForm.Items.Item("13").Specific;

            SAPbouiCOM.ComboBox cbCopyTo = (SAPbouiCOM.ComboBox)this.UIAPIRawForm.Items.Item("2310000078").Specific;
            SAPbouiCOM.ComboBox cbCopyFrom = (SAPbouiCOM.ComboBox)this.UIAPIRawForm.Items.Item("2310000079").Specific;

            SAPbouiCOM.StaticText stDocRef = (SAPbouiCOM.StaticText)this.UIAPIRawForm.Items.Item("256000001").Specific;
            SAPbouiCOM.Button btDocRef = (SAPbouiCOM.Button)this.UIAPIRawForm.Items.Item("256000002").Specific;
            SAPbouiCOM.StaticText stDocRef2 = (SAPbouiCOM.StaticText)this.UIAPIRawForm.Items.Item("234000001").Specific;
            

            SAPbouiCOM.EditText edtRef2 = (SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("21").Specific;
            SAPbouiCOM.StaticText stRef2 = (SAPbouiCOM.StaticText)this.UIAPIRawForm.Items.Item("22").Specific;
            SAPbouiCOM.Matrix mtxAttachments = (SAPbouiCOM.Matrix)this.UIAPIRawForm.Items.Item("1320000081").Specific;

            SAPbouiCOM.ComboBox cbBrowse = (SAPbouiCOM.ComboBox)this.UIAPIRawForm.Items.Item("1320000082").Specific;
            cbBrowse.Item.Top = mtxAttachments.Item.Top + 11;

            SAPbouiCOM.Button btDisplay = (SAPbouiCOM.Button)this.UIAPIRawForm.Items.Item("1320000083").Specific;
            btDisplay.Item.Top = cbBrowse.Item.Top + cbBrowse.Item.Height + 15;


            SAPbouiCOM.Folder fdContents = (SAPbouiCOM.Folder)this.UIAPIRawForm.Items.Item("1320000079").Specific;
            SAPbouiCOM.Folder fdAttach = (SAPbouiCOM.Folder)this.UIAPIRawForm.Items.Item("1320000080").Specific;

            fdContents.Item.Top = edtWONum.Item.Top + 45;
            fdAttach.Item.Top = edtWONum.Item.Top + 45;


            mtxContents.Item.Top = fdContents.Item.Top + 25;
            mtxAttachments.Item.Top = fdAttach.Item.Top + 25;

            
            

            SAPbouiCOM.Item rect1 = (SAPbouiCOM.Item)this.UIAPIRawForm.Items.Item("1320000085");
            SAPbouiCOM.Item rect2 = (SAPbouiCOM.Item)this.UIAPIRawForm.Items.Item("1320000086");
            SAPbouiCOM.Item rect3 = (SAPbouiCOM.Item)this.UIAPIRawForm.Items.Item("1320000087");
            SAPbouiCOM.Item rect4 = (SAPbouiCOM.Item)this.UIAPIRawForm.Items.Item("1320000088");

            rect1.Top = fdContents.Item.Top + 19;
            rect1.Height = (edtRemarks.Item.Top - 10) - rect1.Top;
            rect2.Top = fdContents.Item.Top + 19;
            rect2.Height = (edtRemarks.Item.Top - 10) - rect2.Top;
            rect3.Top = fdContents.Item.Top + 19;
            rect3.Height = (edtRemarks.Item.Top - 10) - rect3.Top;
            rect4.Top = fdContents.Item.Top + 19;
            rect4.Height = (edtRemarks.Item.Top - 10) - rect4.Top;


            mtxContents.Item.Height = (edtRemarks.Item.Top - 20) - mtxContents.Item.Top;
            mtxAttachments.Item.Height = (edtRemarks.Item.Top - 20) - mtxAttachments.Item.Top;

            edtRef2.Item.Top = edtRemarks.Item.Top;
            edtRef2.Item.Left = cbCopyTo.Item.Left;

            stRef2.Item.Top = edtRemarks.Item.Top;
            stRef2.Item.Left = cbCopyFrom.Item.Left + 11;

            stDocRef.Item.Top = edtRemarks.Item.Top + stRef2.Item.Height + 2;
            stDocRef.Item.Left = cbCopyFrom.Item.Left + 11;
            btDocRef.Item.Top = edtRemarks.Item.Top + stRef2.Item.Height + 2;
            btDocRef.Item.Left = stDocRef.Item.Left + stDocRef.Item.Width + 1;

            stDocRef2.Item.Top = edtRemarks.Item.Top + stRef2.Item.Height + 2;
            stDocRef2.Item.Left = btDocRef.Item.Left + btDocRef.Item.Width + 1;

            this.UIAPIRawForm.Freeze(false);

        }

        private Button Button2;
        private ToolFind formToolInfo;
    }
}
