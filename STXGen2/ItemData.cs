using System;
using System.Collections.Generic;
using SAPbouiCOM;

namespace STXGen2
{
    internal class ItemData
    {
        public string ItemCode { get; set; }
        public string U_STXCC1 { get; set; }
        public string U_STXCC2 { get; set; }

        internal static List<ItemData> ConvertDataTableToList(DataTable selectedDataTable)
        {
            List<ItemData> itemsList = new List<ItemData>();
            if (selectedDataTable == null || selectedDataTable.Rows.Count < 1)
            {
                return itemsList;
            }
            for (int i = 0; i < selectedDataTable.Rows.Count; i++)
            {
                ItemData item = new ItemData
                {
                    ItemCode = selectedDataTable.GetValue("ItemCode", i).ToString(),
                    U_STXCC1 = selectedDataTable.GetValue("U_STXCC1", i).ToString(),
                    U_STXCC2 = selectedDataTable.GetValue("U_STXCC2", i).ToString(),
                };
                itemsList.Add(item);
            }

            return itemsList;
        }

        internal static void DisableNBOinfo(IForm UIAPIRawForm)
        {
            SAPbouiCOM.Item eMkseg1 = UIAPIRawForm.Items.Item("MKSeg1");
            eMkseg1.Enabled = false;
            SAPbouiCOM.Item eBrand = UIAPIRawForm.Items.Item("STXBrand");
            eBrand.Enabled = false;
        }

        internal static void EnableNBOinfo(IForm UIAPIRawForm)
        {
            SAPbouiCOM.Item eMkseg1 = UIAPIRawForm.Items.Item("MKSeg1");
            eMkseg1.Enabled = true;
            SAPbouiCOM.Item eBrand = UIAPIRawForm.Items.Item("STXBrand");
            eBrand.Enabled = true;
        }

        internal static void QCIDColumnsDisable(IForm UIAPIRawForm)
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)UIAPIRawForm.Items.Item("38").Specific;
            SAPbouiCOM.Column oColumn = oMatrix.Columns.Item("U_STXQC19ID");
            oColumn.Editable = false;
            oColumn = oMatrix.Columns.Item("U_STXToolNum");
            oColumn.Editable = false;
            oColumn = oMatrix.Columns.Item("U_STXPartNum");
            oColumn.Editable = false;
            oColumn = oMatrix.Columns.Item("U_STXPartName");
            oColumn.Editable = false;
            oColumn = oMatrix.Columns.Item("U_STXLeadTime");
            oColumn.Editable = false;
        }

        internal static void QCIDColumnsEnable(IForm UIAPIRawForm)
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)UIAPIRawForm.Items.Item("38").Specific;
            SAPbouiCOM.Column oColumn = oMatrix.Columns.Item("U_STXQC19ID");
            oColumn.Editable = true;
            oColumn = oMatrix.Columns.Item("U_STXToolNum");
            oColumn.Editable = true;
            oColumn = oMatrix.Columns.Item("U_STXPartNum");
            oColumn.Editable = true;
            oColumn = oMatrix.Columns.Item("U_STXPartName");
            oColumn.Editable = true;
            oColumn = oMatrix.Columns.Item("U_STXLeadTime");
            oColumn.Editable = true;
        }
    }
}