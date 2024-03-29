﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using SAPbouiCOM;
using STXGen2.Properties;
using System.Xml;
using SAPbobsCOM;

namespace STXGen2
{
    internal class QCEvents
    {
        public static HashSet<int> deletedRows = new HashSet<int>();
        public readonly object MatrixLock = new object();
        public static bool _processChooseFromList = false;
        public static Dictionary<int, string> _pendingCFLUpdates = new Dictionary<int, string>();
        private static string xmlOperations = "";

        public static string defValue { get; set; }
        public static object QCLength { get; private set; }
        public static int processOperationsListErr { get; private set; } = 0;
        public static bool operationsUpdate { get; set; } = false;
        public static SAPbouiCOM.DataTable operations { get; set; }

         public static void AddLineToTexturesMatrix(SAPbouiCOM.Form oForm, SAPbouiCOM.Matrix texturesMatrix, int selectedRow)
        {
            oForm.Freeze(true);

            SAPbouiCOM.DBDataSource oDBDataSource = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("@STXQC19T");

            if (oDBDataSource.Size == 5)
            {
                Program.SBO_Application.SetStatusBarMessage("Maximum number of textures reached.", BoMessageTime.bmt_Medium, false);
                oForm.Freeze(false);
                return;
            }


            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(oDBDataSource.GetAsXML());
            XmlNodeList rows = xmlDoc.GetElementsByTagName("row");
            XmlElement newRow = xmlDoc.CreateElement("row");
            XmlNode selectedNode = rows.Item(selectedRow);

            XmlElement cells = xmlDoc.CreateElement("cells");
            newRow.AppendChild(cells);

            XmlElement newCell = xmlDoc.CreateElement("cell");
            XmlElement uidElement = xmlDoc.CreateElement("uid");
            uidElement.InnerText = "VisOrder";
            newCell.AppendChild(uidElement);
            XmlElement valueElement = xmlDoc.CreateElement("value");
            valueElement.InnerText = (selectedRow).ToString();
            newCell.AppendChild(valueElement);
            cells.AppendChild(newCell);


            //rows.Item(selectedRow-1).ParentNode.InsertBefore(newRow, rows.Item(selectedRow - 1));
            rows.Item(selectedRow - 1).ParentNode.AppendChild(newRow);
            for (int i = selectedRow; i < rows.Count; i++)
            {
                XmlNode visOrderNode = rows.Item(i).SelectSingleNode("cells/cell[uid='VisOrder']/value");
                if (visOrderNode != null)
                {
                    visOrderNode.InnerText = (i + 1).ToString();
                }
            }

            string xmlData = xmlDoc.OuterXml;
            oDBDataSource.LoadFromXML(xmlData);

            if (!oForm.Mode.Equals(BoFormMode.fm_UPDATE_MODE))
            {
                oForm.Mode = BoFormMode.fm_UPDATE_MODE;
            }
            oForm.Freeze(false);
        }


        public static void AddLineToOperationMatrix(SAPbouiCOM.Form oForm, Matrix operationsMatrix, int selectedRow)
        {
            bool confirmTOper = false;
            oForm.Freeze(true);

            SAPbouiCOM.DBDataSource oDBDataSource = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("@STXQC19O");


            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(oDBDataSource.GetAsXML());
            XmlNodeList rows = xmlDoc.GetElementsByTagName("row");
            XmlElement newRow = xmlDoc.CreateElement("row");
            XmlNode selectedNode = rows.Item(selectedRow - 1);

            XmlNode cellValueNode = selectedNode.SelectSingleNode("cells/cell[uid='U_seq']/value");
            string OP_seq = cellValueNode?.InnerText;


            cellValueNode = selectedNode.SelectSingleNode("cells/cell[uid='U_Texture']/value");
            string OP_Texture = cellValueNode?.InnerText;

            if (!string.IsNullOrEmpty(OP_Texture))
            {
                confirmTOper = Program.SBO_Application.MessageBox("Is the new line related to the texture of the select line?", 1, "Yes", "No") == 1;
            }

            XmlElement cells = xmlDoc.CreateElement("cells");
            newRow.AppendChild(cells);

            XmlElement newCell = xmlDoc.CreateElement("cell");
            XmlElement uidElement = xmlDoc.CreateElement("uid");
            uidElement.InnerText = "VisOrder";
            newCell.AppendChild(uidElement);
            XmlElement valueElement = xmlDoc.CreateElement("value");
            valueElement.InnerText = (selectedRow).ToString();
            newCell.AppendChild(valueElement);
            cells.AppendChild(newCell);

            if (confirmTOper)
            {
                newCell = xmlDoc.CreateElement("cell");
                uidElement = xmlDoc.CreateElement("uid");
                uidElement.InnerText = "U_Texture";
                newCell.AppendChild(uidElement);
                valueElement = xmlDoc.CreateElement("value");
                valueElement.InnerText = (OP_Texture).ToString();
                newCell.AppendChild(valueElement);
                cells.AppendChild(newCell);

                newCell = xmlDoc.CreateElement("cell");
                uidElement = xmlDoc.CreateElement("uid");
                uidElement.InnerText = "U_seq";
                newCell.AppendChild(uidElement);
                valueElement = xmlDoc.CreateElement("value");
                valueElement.InnerText = (OP_seq).ToString();
                newCell.AppendChild(valueElement);
                cells.AppendChild(newCell);
            }

            if (selectedRow == rows.Count) // When the last row is selected
            {
                // Append the new row to the end of the list
                rows.Item(selectedRow-1).ParentNode.AppendChild(newRow);
            }
            else
            {
                // Insert the new row before the next row
                rows.Item(selectedRow).ParentNode.InsertBefore(newRow, rows.Item(selectedRow - 1));
            }
            //rows.Item(selectedRow).ParentNode.InsertBefore(newRow, rows.Item(selectedRow - 1));
            for (int i = selectedRow; i < rows.Count; i++)
            {
                XmlNode visOrderNode = rows.Item(i).SelectSingleNode("cells/cell[uid='VisOrder']/value");
                if (visOrderNode != null)
                {
                    visOrderNode.InnerText = (i + 1).ToString();
                }
            }

            string xmlData = xmlDoc.OuterXml;
            oDBDataSource.LoadFromXML(xmlData);

            if (!oForm.Mode.Equals(BoFormMode.fm_UPDATE_MODE))
            {
                oForm.Mode = BoFormMode.fm_UPDATE_MODE;
            }
            oForm.Freeze(false);
        }

        internal static void FillTextureClass(IForm uIAPIRawForm)
        {
            SAPbouiCOM.Form oForm = ((SAPbouiCOM.Form)uIAPIRawForm);

            // Get a reference to the matrix object
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("mTextures").Specific;

            // Check if mTextures matrix has 0 rows and add a new row if needed
            if (oMatrix.RowCount == 0)
            {
                oMatrix.AddRow();
            }

            // Get a reference to the existing column
            SAPbouiCOM.Column oColumn = oMatrix.Columns.Item("QCTClass");

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Utils.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string sQuery = "SELECT Code, Name FROM \"@STXQC19TCLASS\"";
            oRecordSet.DoQuery(sQuery);

            while (!oRecordSet.EoF)
            {
                string sCode = oRecordSet.Fields.Item("Code").Value.ToString();
                string sName = oRecordSet.Fields.Item("Name").Value.ToString();

                // Iterate through the rows of the matrix and set the ComboBox values
                for (int i = 1; i <= oMatrix.RowCount; i++)
                {
                    SAPbouiCOM.ComboBox oComboBox = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item("QCTClass").Cells.Item(i).Specific;
                    oComboBox.ValidValues.Add(sCode, sName);
                }

                oRecordSet.MoveNext();
            }
        }

        internal static void FillUnitMeasures(IForm uIAPIRawForm)
        {
            string sUOM = "";
            SAPbouiCOM.Form oForm = ((SAPbouiCOM.Form)uIAPIRawForm);
            SAPbouiCOM.ComboBox UnMsr = (SAPbouiCOM.ComboBox)oForm.Items.Item("UnMsr").Specific;
            SAPbouiCOM.EditText QCLength = (SAPbouiCOM.EditText)oForm.Items.Item("QCLength").Specific;
            SAPbouiCOM.EditText QCDocEntry = (SAPbouiCOM.EditText)oForm.Items.Item("QCDocEntry").Specific;

            string s = QCDocEntry.Value.ToString();

            string query = "select \"UnitDisply\",\"UnitName\",(select T0.\"UnitDisply\" from \"OLGT\" T0 inner join \"OADM\" T1 on T0.\"UnitCode\" = T1.\"DefLengthU\") as \"DefLengthU\" from \"OLGT\"";
            SAPbobsCOM.Recordset recordset = (SAPbobsCOM.Recordset)Utils.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            recordset.DoQuery(query);
            while (!recordset.EoF)
            {
                string sCode = recordset.Fields.Item("UnitDisply").Value.ToString();
                string sName = recordset.Fields.Item("UnitName").Value.ToString();
                defValue = recordset.Fields.Item("DefLengthU").Value.ToString();
                UnMsr.ValidValues.Add(sCode, sName);
                recordset.MoveNext();
            }

            string query2 = $"select coalesce(\"U_pLength\",\"U_pWidth\") as \"InitialUom\" from \"@STXQC19\" where \"DocEntry\" = '{SAPEvents.qcid}'";
            SAPbobsCOM.Recordset recordset2 = (SAPbobsCOM.Recordset)Utils.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            recordset2.DoQuery(query2);
            while (!recordset2.EoF)
            {

                sUOM = recordset2.Fields.Item("InitialUom").Value.ToString();
                sUOM = sUOM.Substring(sUOM.IndexOf(" ") + 1);
                recordset2.MoveNext();
            }


            if (string.IsNullOrEmpty(sUOM))
            {
                UnMsr.Select(defValue, SAPbouiCOM.BoSearchKey.psk_ByValue);
            }
            else
            {
                UnMsr.Select(sUOM, SAPbouiCOM.BoSearchKey.psk_ByValue);
            }

        }

        internal static void UpdateCovArea(IForm uIAPIRawForm, string previousUOM, string selectedUOM, bool isUoMAreaChanging)
        {
            System.Globalization.NumberFormatInfo sapNumberFormat = Utils.GetSAPNumberFormatInfo();

            SAPbouiCOM.Form oForm = ((SAPbouiCOM.Form)uIAPIRawForm);
            SAPbouiCOM.Matrix mTextures = (SAPbouiCOM.Matrix)oForm.Items.Item("mTextures").Specific;

            // Loop through each row in the mtxQCItems Matrix control
            for (int i = 1; i <= mTextures.RowCount; i++)
            {
                SAPbouiCOM.EditText selTexture = (SAPbouiCOM.EditText)mTextures.Columns.Item("QCTexture").Cells.Item(i).Specific;
                if (!string.IsNullOrEmpty(selTexture.Value))
                {
                    if (isUoMAreaChanging == true)
                    {
                        SAPbouiCOM.EditText CovArea = (SAPbouiCOM.EditText)mTextures.Columns.Item("QCCovA").Cells.Item(i).Specific;
                        SAPbouiCOM.EditText calArea = (SAPbouiCOM.EditText)oForm.Items.Item("QCArea").Specific;

                        mTextures.SetCellWithoutValidation(i, "QCCovA", calArea.Value);
                    }
                    else
                    {
                        SAPbouiCOM.EditText CovArea = (SAPbouiCOM.EditText)mTextures.Columns.Item("QCCovA").Cells.Item(i).Specific;

                        double covA = HelperMethods.ParseDoubleWUOM(CovArea.Value, sapNumberFormat); 
                        double convertedcovA = DBCalls.ConvertAreaDimensions(covA, selectedUOM, previousUOM);
                        string formattedValue = convertedcovA.ToString("N", sapNumberFormat);
                        mTextures.SetCellWithoutValidation(i, "QCCovA", $"{formattedValue} {selectedUOM}²");
                    }
                }
            }
        }

        internal static string CalculateArea(string formUID, string selectedUoM)
        {
            System.Globalization.NumberFormatInfo sapNumberFormat = Utils.GetSAPNumberFormatInfo();

            SAPbouiCOM.Form oForm = Program.SBO_Application.Forms.Item(formUID);
            EditText edtQCLength = (EditText)oForm.Items.Item("QCLength").Specific;
            EditText edtQCWidth = (EditText)oForm.Items.Item("QCWidth").Specific;
            EditText edtQCArea = (EditText)oForm.Items.Item("QCArea").Specific;

            double length = HelperMethods.ParseDoubleWUOM(edtQCLength.Value, sapNumberFormat);
            double width = HelperMethods.ParseDoubleWUOM(edtQCWidth.Value, sapNumberFormat);

            double area = length * width;
            string areaFormatted = area.ToString("N", sapNumberFormat);
            edtQCArea.Value = $"{areaFormatted} {selectedUoM}²";

            return edtQCArea.Value;
        }

        internal static string SellMarginImage(IForm uIAPIRawForm)
        {
            System.Globalization.NumberFormatInfo sapNumberFormat = Utils.GetSAPNumberFormatInfo();
            string resourceName = "";
            double compPrice = 0;
            double compCost = 0;
            string DocCur = ((SAPbouiCOM.EditText)uIAPIRawForm.Items.Item("QCDocCur").Specific).Value;
            string Cost = ((SAPbouiCOM.EditText)uIAPIRawForm.Items.Item("QCTEst").Specific).Value;
            string Price = ((SAPbouiCOM.EditText)uIAPIRawForm.Items.Item("UnPrice").Specific).Value;
            string LCPrice = ((SAPbouiCOM.EditText)uIAPIRawForm.Items.Item("LCPrice").Specific).Value;

            if (Utils.MainCurrency != DocCur)
            {
                compPrice = HelperMethods.ParseDoubleWCur(LCPrice, sapNumberFormat);

            }
            else
            {
                compPrice = HelperMethods.ParseDoubleWCur(Price, sapNumberFormat);
            }

            compCost = double.Parse(Cost, System.Globalization.NumberStyles.AllowDecimalPoint | System.Globalization.NumberStyles.AllowThousands, sapNumberFormat);

            
            // Your condition to choose the image
            if (compCost < compPrice)
            {
                resourceName = "Light-green.jpg";
            }
            else if (compCost == compPrice)
            {
                resourceName = "Light-yellow.jpg";
            }
            else
            {
                resourceName = "Light-red.jpg";
            }

            string imagePath = HelperMethods.GetAndSaveImage(resourceName);
            return imagePath;

        }

        internal static List<Dictionary<string, string>> GetAllValuesFromMatrix1(Matrix mTextures)
        {
            // Get the number of rows and columns in Matrix1
            int rowCount = mTextures.RowCount;
            int columnCount = mTextures.Columns.Count;

            // Create a list to store the matrix values
            List<Dictionary<string, string>> mTexturesValues = new List<Dictionary<string, string>>();

            // Loop through the rows of Matrix1
            for (int rowIndex = 1; rowIndex <= rowCount; rowIndex++)
            {
                // Create a dictionary to store the row values
                Dictionary<string, string> rowValues = new Dictionary<string, string>();

                // Get the values of the TextureCode, Quantity, and CoverageArea columns
                string textureCode = ((SAPbouiCOM.EditText)mTextures.Columns.Item("QCTexture").Cells.Item(rowIndex).Specific).Value;
                string quantity = ((SAPbouiCOM.EditText)mTextures.Columns.Item("QCQuantity").Cells.Item(rowIndex).Specific).Value;
                string coverageArea = ((SAPbouiCOM.EditText)mTextures.Columns.Item("QCCovA").Cells.Item(rowIndex).Specific).Value;
                string TClass = ((SAPbouiCOM.ComboBox)mTextures.Columns.Item("QCTClass").Cells.Item(rowIndex).Specific).Value;
                string GComplex = ((SAPbouiCOM.ComboBox)mTextures.Columns.Item("QCGComp").Cells.Item(rowIndex).Specific).Value;

                if (!string.IsNullOrEmpty(textureCode))
                {
                    // Store the values in the row dictionary
                    rowValues["QCTexture"] = textureCode;
                    rowValues["QCQuantity"] = quantity;
                    rowValues["QCCovA"] = coverageArea;
                    rowValues["QCTClass"] = TClass;
                    rowValues["QCGComp"] = GComplex;


                    // Add the row dictionary to the matrix values list
                    mTexturesValues.Add(rowValues);
                }
            }

            return mTexturesValues;
        }

        internal static void RemoveLinefromOperationMatrix(SAPbouiCOM.Form oForm, Matrix OperationsMatrix, int selectedMatrixRow)
        {
            try
            {
                oForm.Freeze(true);
                SAPbouiCOM.ComboBox OPFilter = (SAPbouiCOM.ComboBox)oForm.Items.Item("OPFilter").Specific;
                if (SAPEvents.selectedRow > 0)
                {
                    SAPbouiCOM.DBDataSource oDBDataSource = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("@STXQC19O");

                    // Remove the row from the data source
                    oDBDataSource.RemoveRecord(SAPEvents.selectedRow - 1);

                    // Update the # aka LineID column
                    for (int i = SAPEvents.selectedRow - 1; i < oDBDataSource.Size; i++)
                    {
                        oDBDataSource.SetValue("VisOrder", i, (i + 1).ToString());
                    }

                    // Refresh the matrix
                    OperationsMatrix.LoadFromDataSource();

                    // Set this property to prevent the selection of the row after deletion.
                    OperationsMatrix.SelectionMode = BoMatrixSelect.ms_None;

                    // Set the selection mode back to the default after loading the data source.
                    OperationsMatrix.SelectionMode = BoMatrixSelect.ms_Auto;
                    QCEvents.OperationsCalcTotal(oForm);
                }
            }
            catch (Exception)
            {
            }
            finally
            {
                oForm.Freeze(false);
            }
            
        }

        internal static void RemoveLinefromTexturesMatrix(SAPbouiCOM.Form oForm, Matrix texturesMatrix, int selectedMatrixRow)
        {
            try
            {
                oForm.Freeze(true);
                if (SAPEvents.selectedRow > 0)
                {
                    SAPbouiCOM.DBDataSource oDBDataSource = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("@STXQC19T");

                    // Remove the row from the data source
                    oDBDataSource.RemoveRecord(SAPEvents.selectedRow - 1);

                    // Update the # aka LineID column
                    for (int i = SAPEvents.selectedRow - 1; i < oDBDataSource.Size; i++)
                    {
                        oDBDataSource.SetValue("VisOrder", i, (i + 1).ToString());
                    }

                    // Refresh the matrix
                    texturesMatrix.LoadFromDataSource();

                    // Set this property to prevent the selection of the row after deletion.
                    texturesMatrix.SelectionMode = BoMatrixSelect.ms_None;

                    // Set the selection mode back to the default after loading the data source.
                    texturesMatrix.SelectionMode = BoMatrixSelect.ms_Auto;
                }
            }
            catch (Exception)
            {

            }
            finally
            {
                oForm.Freeze(false);
            }
           
        }

        internal static void GetSubPartType(IForm uIAPIRawForm, SAPbouiCOM.EditText qCSubPart)
        {
            uIAPIRawForm.Freeze(true);
            string spt = "";
            string descr = "";
            SAPbouiCOM.Form oForm = ((SAPbouiCOM.Form)uIAPIRawForm);
            SAPbouiCOM.EditText SubPartType = (SAPbouiCOM.EditText)oForm.Items.Item("QCPartType").Specific;
            SAPbouiCOM.EditText PartDescr = (SAPbouiCOM.EditText)oForm.Items.Item("SPartDescr").Specific;


            string checkUDSId = $"CheckSPT";
            if (!uIAPIRawForm.DataSources.UserDataSources.Cast<SAPbouiCOM.UserDataSource>().Any(uds => uds.UID == checkUDSId))
            {
                uIAPIRawForm.DataSources.UserDataSources.Add(checkUDSId, BoDataType.dt_SHORT_TEXT, 100);
            }

            // Get the user data source
            SAPbouiCOM.UserDataSource oDS = oForm.DataSources.UserDataSources.Item(checkUDSId);

            // Bind the SubPartType field to the user data source
            SubPartType.DataBind.SetBound(true, "", checkUDSId);

            (spt, descr) = DBCalls.GetSPT(qCSubPart);

            // Update the value via the user data source
            oDS.Value = spt;
            //QuoteCalculator.parttDescr = descr;
            PartDescr.Value = descr;

            uIAPIRawForm.Freeze(false);
        }

   
        internal static (string AdditionalConditions, string ConcatenatedTextureCodes, string tClassExpression, string OpQuantityExpression, string QtyFactorExpression) GetAdditionalConditions(List<Dictionary<string, string>> matrix1Values)
        {
            string quantity = "";
            string textureCode = "";
            string coverageArea = "0";
            string tClass = "";
            string GeoComplex = "";

            var calcFactorList = new List<string>();
            var TextureClassList = new List<string>();
            var OpQuantityList = new List<string>();
            var QtycFactorList = new List<string>();
            string concatenatedTextureCodes = GetConcatenatedTextureCodes(matrix1Values);


            for (int i = 0; i < matrix1Values.Count; i++)
            {
                Dictionary<string, string> rowValues = matrix1Values[i];

                textureCode = rowValues["QCTexture"];
                quantity = rowValues["QCQuantity"];
                tClass = rowValues["QCTClass"];
                GeoComplex = rowValues["QCGComp"];
                coverageArea = Regex.Replace(rowValues["QCCovA"], $@"[^\d{Utils.decSep}{Utils.thousSep}]", "");
                coverageArea = DBCalls.ConvertDimMeters(HelperMethods.ParseSAPValueToDouble(coverageArea), QuoteCalculator.selectedUOM);

                string condition1 = $"WHEN T2.\"U_standexReference\" = '{textureCode}' AND T1.\"U_STXQtyBy\" = 'A' THEN {coverageArea}";
                calcFactorList.Add(condition1);

                string condition2 = $"WHEN T2.\"U_standexReference\" = '{textureCode}' THEN (select \"U_Factor\" from \"@STXQC19TCLASS\" Where \"Code\" = {tClass})";
                TextureClassList.Add(condition2);

                string condition3 = $"WHEN {GeoComplex} = 1 then T1.\"Quantity\" when {GeoComplex} = 2 then T1.\"U_STXQTYGC2\" when {GeoComplex} = 3 then T1.\"U_STXQTYGC3\"";
                OpQuantityList.Add(condition3);

                string condition4 = $"WHEN T2.\"U_standexReference\" = '{textureCode}' THEN {quantity}";
                QtycFactorList.Add(condition4);



            }

            string calcFactorcond = string.Join(" ", calcFactorList);
            string calcFactorExpression = $"(CASE {calcFactorcond} ELSE 1 END) as \"CalcFactor\"";

            string tClasscond = string.Join(" ", TextureClassList);
            string tClassExpression = $"(CASE {tClasscond} ELSE 1 END) as \"TClassFactor\"";

            string OPQty = string.Join(" ", OpQuantityList);
            string OpQuantityExpression = $"(CASE {OPQty} ELSE 0 END) as \"Quantity\"";

            string QtyFactor = string.Join(" ", QtycFactorList);
            string QtyFactorExpression = $"(Case {QtyFactor} ELSE 1 END) as \"QtyFactor\"";

            return (calcFactorExpression, concatenatedTextureCodes, tClassExpression, OpQuantityExpression, QtyFactorExpression);

        }

        internal static void GetFiltersOperations(IForm uIAPIRawForm, EditText qCDocEntry)
        {
            SAPbouiCOM.ComboBox comboBox = (SAPbouiCOM.ComboBox)uIAPIRawForm.Items.Item("OPFilter").Specific;

            // Clear existing values
            while (comboBox.ValidValues.Count > 0)
            {
                comboBox.ValidValues.Remove(0, BoSearchKey.psk_Index);
            }

            DBCalls.GetFilterOperations(comboBox, qCDocEntry);
           
            comboBox.Select(0, BoSearchKey.psk_Index);
        }


        private static string GetConcatenatedTextureCodes(List<Dictionary<string, string>> matrix1Values)
        {
            var textureCodes = matrix1Values.Select(x => $"'{x["QCTexture"]}'");
            return string.Join(",", textureCodes);
        }

        internal static void GetDefOperations(IForm uIAPIRawForm, int selectedRow)
        {
            processOperationsListErr = 0;

            SAPbouiCOM.Matrix mOperations = (Matrix)uIAPIRawForm.Items.Item("mOper").Specific;
            SAPbouiCOM.Matrix matrix1 = (Matrix)uIAPIRawForm.Items.Item("mTextures").Specific;

            List<Dictionary<string, string>> matrix1Values = QCEvents.GetAllValuesFromMatrix1(matrix1);

            processOperationsList(uIAPIRawForm, matrix1Values);
            processMTOperationsList(uIAPIRawForm, mOperations, matrix1Values, selectedRow);
        }

        internal static void GetOperations(IForm uIAPIRawForm, int selectedRow)
        {
            processOperationsListErr = 0;

            SAPbouiCOM.Matrix mOperations = (Matrix)uIAPIRawForm.Items.Item("mOper").Specific;
            SAPbouiCOM.Matrix matrix1 = (Matrix)uIAPIRawForm.Items.Item("mTextures").Specific;

            List<Dictionary<string, string>> matrix1Values = QCEvents.GetAllValuesFromMatrix1(matrix1);

            processOperationsList(uIAPIRawForm, matrix1Values);

            switch (processOperationsListErr)
            {
                case 0:
                    processMTOperationsList(uIAPIRawForm, mOperations, matrix1Values, selectedRow);
                    break;
                case 1:
                    Program.SBO_Application.SetStatusBarMessage("Selection of SPT missing.", BoMessageTime.bmt_Medium, false);
                    break;
                case 2:
                    Program.SBO_Application.SetStatusBarMessage("No textures selected.", BoMessageTime.bmt_Medium, false);
                    break;
                default:
                    break;
            }
        }

        private static void processMTOperationsList(IForm uIAPIRawForm, Matrix mOperations, List<Dictionary<string, string>> mtTexture, int selRow)
        {
            SAPbouiCOM.DBDataSource oDBDataSource = (SAPbouiCOM.DBDataSource)uIAPIRawForm.DataSources.DBDataSources.Item("@STXQC19O");


            var filteredOperations = QCEvents.GetFilteredOperations(uIAPIRawForm, selRow);

            var (CalcFactorConditions, concatenatedTextureCodes, tclassConditions, OpQuantityExpression, QtyFactorExpression) = QCEvents.GetAdditionalConditions(mtTexture);

            // Create a unique identifier for the DataTable
            string dataTableID = "Operations";

            // Check if the DataTable with the ID "Operations" exists
            if (!DataTableExists(uIAPIRawForm, dataTableID))
            {
                operations = uIAPIRawForm.DataSources.DataTables.Add(dataTableID);
            }
            else
            {
                operations = uIAPIRawForm.DataSources.DataTables.Item(dataTableID);
            }

            if (((SAPbouiCOM.CheckBox)uIAPIRawForm.Items.Item("DefBOM").Specific).Checked == true)
            {
                xmlOperations = DBCalls.GetOperation(operations, uIAPIRawForm, mOperations, CalcFactorConditions, concatenatedTextureCodes, tclassConditions, OpQuantityExpression, ((SAPbouiCOM.EditText)uIAPIRawForm.Items.Item("QCItemCode").Specific).Value, ((SAPbouiCOM.CheckBox)uIAPIRawForm.Items.Item("DefBOM").Specific).Checked, QtyFactorExpression, filteredOperations);
            }
            else
            {
                xmlOperations = DBCalls.GetOperation(operations, uIAPIRawForm, mOperations, CalcFactorConditions, concatenatedTextureCodes, tclassConditions, OpQuantityExpression, ((SAPbouiCOM.EditText)uIAPIRawForm.Items.Item("QCSubPart").Specific).Value, ((SAPbouiCOM.CheckBox)uIAPIRawForm.Items.Item("DefBOM").Specific).Checked, QtyFactorExpression, filteredOperations);
            }

            int operationscount = operations.Rows.Count;
            if (operationscount > 0)
            {
                try
                {
                    uIAPIRawForm.Freeze(true);

                    mOperations.Clear();
                    
                    oDBDataSource.LoadFromXML(xmlOperations);

                    // Bind check boxes using UserDataSources to be able to multiselect
                    BindMatrixCheckboxes(uIAPIRawForm, mOperations, operationscount);

                    // Load data from the DataTable to the matrix
                    mOperations.LoadFromDataSource();

                    mOperations.AutoResizeColumns();

                    // Color all rows with Error message on the matrix
                    SetMatrixRowColor(mOperations, operations, "OPErrMsg");

                }
                catch (Exception ex)
                {
                    Program.SBO_Application.SetStatusBarMessage(ex.Message, BoMessageTime.bmt_Medium, true);
                }
                finally
                {
                    uIAPIRawForm.Freeze(false);
                    operationsUpdate = true;

                    uIAPIRawForm.Mode = BoFormMode.fm_UPDATE_MODE;

                    Program.SBO_Application.SetStatusBarMessage("Operations matrix updated.", BoMessageTime.bmt_Short, false);
                }
            }
        }

        private static string GetFilteredOperations(IForm uIAPIRawForm,int selRow)
        {

            SAPbouiCOM.Form parentForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.Item(Utils.ParentFormUID);
            int objectType = parentForm.Type;
            BoObjectTypes docobjtype = DBCalls.GetSAPObjectType(objectType.ToString());
            string strObjType = DBCalls.GetSAPObjectLineStr(docobjtype);

            SAPbouiCOM.Item docNumItem = parentForm.Items.Item("8");
            SAPbouiCOM.EditText docNumEditText = (SAPbouiCOM.EditText)docNumItem.Specific;
            string docNumber = docNumEditText.Value;

            SAPbouiCOM.Matrix parentMatrix = (SAPbouiCOM.Matrix)parentForm.Items.Item("38").Specific;
            SAPbouiCOM.EditText docLine = (SAPbouiCOM.EditText)parentMatrix.Columns.Item("110").Cells.Item(selRow).Specific;

            string mainStrObjType = strObjType.Substring(0, 3);

            SAPbouiCOM.EditText toolnum = (SAPbouiCOM.EditText)uIAPIRawForm.Items.Item("QCToolNum").Specific;

            string calcFactorExpression = $"where X0.\"U_operationCode\" not in \n" +
                                        $"(select \"U_opCode\" from \"@STXQC19O\" T0\n" +
                                        $"inner join \"@STXOPERATIONS\" T1 on T0.\"U_opCode\" = T1.\"Code\"\n" +
                                        $"where T0.\"DocEntry\" in (\n" +
                                        $"SELECT T1.\"U_STXQC19ID\"\n" +
                                        $"FROM O{mainStrObjType} T0\n" +
                                        $"inner join {strObjType} T1 on T0.\"DocEntry\" = T1.\"DocEntry\"\n" +
                                        $"where \"DocNum\" = {docNumber} and T1.\"LineNum\" < {docLine.Value} and(T1.\"U_STXToolNum\" = '{toolnum.Value}' and coalesce(T1.\"U_STXToolNum\", '') <> '')) and \"U_opCode\" <> 'TC00-07' and T1.\"U_PlanType\" in ('I', 'F'))";


            return calcFactorExpression;
        }


        private static bool DataTableExists(IForm uIAPIRawForm, string dataTableID)
        {
            for (int i = 0; i < uIAPIRawForm.DataSources.DataTables.Count; i++)
            {
                SAPbouiCOM.DataTable dt = uIAPIRawForm.DataSources.DataTables.Item(i);
                if (dt.UniqueID == dataTableID)
                {
                    return true;
                }
            }
            return false;
        }

        private static void BindMatrixColumns(Matrix mOperations, string dataTableID)
        {

            string[] columnsToBind = new[] { "#", "OPTexture", "OPResc", "OPResN", "OPcode", "OPName", "OPNameL", "OPStdT", "OPQtdT", "OPUom", "OPCost", "OPTotal", "OPErrMsg", "OPSeq" };
            foreach (string column in columnsToBind)
            {
                    mOperations.Columns.Item(column).DataBind.Bind(dataTableID, column);
            }
        }

        public static void BindMatrixCheckboxes(IForm uIAPIRawForm, Matrix mOperations, int operationscount)
        {
            string checkUDSId = $"CheckUDS";
            if (!uIAPIRawForm.DataSources.UserDataSources.Cast<SAPbouiCOM.UserDataSource>().Any(uds => uds.UID == checkUDSId))
            {
                uIAPIRawForm.DataSources.UserDataSources.Add(checkUDSId, BoDataType.dt_SHORT_TEXT, 1);
            }
            mOperations.Columns.Item("OPcheck").DataBind.SetBound(true, "", checkUDSId);
        }

        private static void SetMatrixRowColor(SAPbouiCOM.Matrix mOperations, SAPbouiCOM.DataTable operations, string colUID)
        {
            Color orangeColor = Color.FromArgb(0xFF, 0xD1, 0x55);
            int warning = (orangeColor.R) + (orangeColor.G << 8) + (orangeColor.B << 16);

            for (int rowIndex = 1; rowIndex <= mOperations.RowCount; rowIndex++)
            {
                string cellValue = operations.GetValue(12, rowIndex - 1).ToString();

                if (!string.IsNullOrEmpty(cellValue))
                {
                    ((SAPbouiCOM.CheckBox)mOperations.Columns.Item("OPcheck").Cells.Item(rowIndex).Specific).Checked = true;
                    mOperations.CommonSetting.SetRowBackColor(rowIndex, warning);
                }
                else
                {
                    mOperations.CommonSetting.SetRowBackColor(rowIndex, -1); // Reset to default color
                }
            }
        }

        private static void processOperationsList(IForm uIAPIRawForm, List<Dictionary<string, string>> matrix1Values)
        {
            SAPbouiCOM.EditText spt = (SAPbouiCOM.EditText)uIAPIRawForm.Items.Item("QCSubPart").Specific;
            if (string.IsNullOrEmpty(spt.Value))
            {
                processOperationsListErr = 1;
            }

            if (matrix1Values.Count == 0)
            {
                processOperationsListErr = 2;
            }

        }

        internal static void GetOperationsGrp(IForm uIAPIRawForm)
        {
            try
            {
                //uIAPIRawForm.Freeze(true);

                SAPbouiCOM.DataTable operations;

                SAPbouiCOM.Matrix mOperations = (Matrix)uIAPIRawForm.Items.Item("mOper").Specific;
                mOperations.Clear();

                SAPbouiCOM.Matrix matrix1 = (Matrix)uIAPIRawForm.Items.Item("mTextures").Specific;
                List<Dictionary<string, string>> matrix1Values = QCEvents.GetAllValuesFromMatrix1(matrix1);


                var (CalcFactorCond, concatenatedTextureCodes, tClassCond, OpQuantityExpression, QtyFactorExpression) = QCEvents.GetAdditionalConditions(matrix1Values);

                try
                {
                    // Try to get the existing DataTable
                    operations = uIAPIRawForm.DataSources.DataTables.Item("Operations");
                    operations.Clear();
                }
                catch
                {
                    // If the DataTable doesn't exist, create a new one
                    operations = uIAPIRawForm.DataSources.DataTables.Add("Operations");
                }
            }
            catch (Exception ex)
            {
                Program.SBO_Application.SetStatusBarMessage(ex.Message, BoMessageTime.bmt_Medium, true);
            }

            //finally
            //{
            //    uIAPIRawForm.Freeze(false);
            //}

        }



        internal static void GetResultsfromFilter(IForm uIAPIRawForm, Matrix mOperations, string selectedValue)
        {

            Color selectionColor = Color.FromArgb(255, 0x83, 0xC5, 0x55);
            int selection = (selectionColor.R) + (selectionColor.G << 8) + (selectionColor.B << 16);

            // Ensure the datasource is updated with the UI values
            mOperations.FlushToDataSource();

            SAPbouiCOM.DBDataSource ds = (SAPbouiCOM.DBDataSource)uIAPIRawForm.DataSources.DBDataSources.Item("@STXQC19O");

            XElement xml = XElement.Parse(ds.GetAsXML());

            var rows = xml.Element("rows").Elements("row").ToList();

            for (int rowIndex = 1; rowIndex <= mOperations.RowCount; rowIndex++)
            {
                var row = rows[rowIndex - 1];
                var cells = row.Element("cells").Elements("cell");

                string opSeqValue = cells.FirstOrDefault(c => c.Element("uid").Value == "U_seq")?.Element("value")?.Value;

                if (selectedValue != "-1" && opSeqValue == selectedValue)
                {
                    mOperations.CommonSetting.SetRowBackColor(rowIndex, selection);
                }
                else
                {
                    mOperations.CommonSetting.SetRowBackColor(rowIndex, -1);
                }
            }

            uIAPIRawForm.Mode = BoFormMode.fm_OK_MODE;
        }



        internal static void OperationsTotalFilter(IForm uIAPIRawForm, string selectedValue)
        {
            double totalop = 0;
            double totalsub = 0;
            double totalqty = 0;

            var mOperations = HelperMethods.GetMatrix(uIAPIRawForm, "mOper");

            if (selectedValue != "-1")
            {
                for (int i = 1; i <= mOperations.RowCount; i++)
                {
                    if (((SAPbouiCOM.EditText)mOperations.Columns.Item("OPSeq").Cells.Item(i).Specific).Value == selectedValue)
                    {
                        var opRescValue = ((SAPbouiCOM.EditText)mOperations.Columns.Item("OPResc").Cells.Item(i).Specific).Value.ToString();
                        if (!opRescValue.StartsWith("SUBCON"))
                        {
                            var optotalCell = (SAPbouiCOM.EditText)mOperations.Columns.Item("OPTotal").Cells.Item(i).Specific;
                            totalop += HelperMethods.ParseValueToDouble(optotalCell.Value);

                            var qtytotalCell = (SAPbouiCOM.EditText)mOperations.Columns.Item("OPQtdT").Cells.Item(i).Specific;
                            totalqty += HelperMethods.ParseValueToDouble(qtytotalCell.Value);
                        }
                        else
                        {
                            var subtotalCell = (SAPbouiCOM.EditText)mOperations.Columns.Item("OPTotal").Cells.Item(i).Specific;
                            totalsub += HelperMethods.ParseValueToDouble(subtotalCell.Value);
                        }
                    }
                }
            }
            else
            {
                totalop = HelperMethods.ParseSAPValueToDouble(((SAPbouiCOM.EditText)uIAPIRawForm.Items.Item("QCOPTot").Specific).Value);
                totalqty = HelperMethods.ParseSAPValueToDouble(((SAPbouiCOM.EditText)uIAPIRawForm.Items.Item("QCTotalH").Specific).Value);
                totalsub = HelperMethods.ParseSAPValueToDouble(((SAPbouiCOM.EditText)uIAPIRawForm.Items.Item("QCTotalSC").Specific).Value);

            }


            HelperMethods.UpdateEditText(uIAPIRawForm, "QCOpA", totalop);
            HelperMethods.UpdateEditText(uIAPIRawForm, "QCTotalSCF", totalsub);
            HelperMethods.UpdateEditText(uIAPIRawForm, "QCTotalHF", totalqty);
        }


        internal static void mtxLineDataRecalculation(IForm uIAPIRawForm, string opResc, EditText opNewQty, string previousQty, string newCost, string previousLineTotal, string itemUID, string previousResc)
        {
            SAPbouiCOM.EditText QCOPTot = (SAPbouiCOM.EditText)uIAPIRawForm.Items.Item("QCOPTot").Specific;
            SAPbouiCOM.EditText QCOpA = (SAPbouiCOM.EditText)uIAPIRawForm.Items.Item("QCOpA").Specific;
            SAPbouiCOM.EditText QCOTCost = (SAPbouiCOM.EditText)uIAPIRawForm.Items.Item("QCOTCost").Specific;
            SAPbouiCOM.EditText QCTEst = (SAPbouiCOM.EditText)uIAPIRawForm.Items.Item("QCTEst").Specific;

            SAPbouiCOM.EditText QCTotalHF = (SAPbouiCOM.EditText)uIAPIRawForm.Items.Item("QCTotalHF").Specific;
            SAPbouiCOM.EditText QCTotalH = (SAPbouiCOM.EditText)uIAPIRawForm.Items.Item("QCTotalH").Specific;

            SAPbouiCOM.EditText QCTotalSCF = (SAPbouiCOM.EditText)uIAPIRawForm.Items.Item("QCTotalSCF").Specific;
            SAPbouiCOM.EditText QCTotalSC = (SAPbouiCOM.EditText)uIAPIRawForm.Items.Item("QCTotalSC").Specific;
            if (itemUID == "mOper")
            {
                if (!opResc.ToString().StartsWith("SUBCON") && !previousResc.ToString().StartsWith("SUBCON"))
                {
                    HelperMethods.UpdateEditText(uIAPIRawForm, "QCOPTot", HelperMethods.ParseSAPValueToDouble(QCOPTot.Value) - HelperMethods.ParseValueToDouble(previousLineTotal) + HelperMethods.ParseValueToDouble(newCost));
                    HelperMethods.UpdateEditText(uIAPIRawForm, "QCOpA", HelperMethods.ParseSAPValueToDouble(QCOpA.Value) - HelperMethods.ParseValueToDouble(previousLineTotal) + HelperMethods.ParseValueToDouble(newCost));
                    //HelperMethods.UpdateEditText(uIAPIRawForm, "QCOTCost", HelperMethods.ParseSAPValueToDouble(QCOTCost.Value) - HelperMethods.ParseValueToDouble(previousLineTotal) + HelperMethods.ParseValueToDouble(newCost));

                    HelperMethods.UpdateEditText(uIAPIRawForm, "QCTotalHF", HelperMethods.ParseSAPValueToDouble(QCTotalHF.Value) - HelperMethods.ParseValueToDouble(previousQty) + HelperMethods.ParseValueToDouble(opNewQty.Value));
                    HelperMethods.UpdateEditText(uIAPIRawForm, "QCTotalH", HelperMethods.ParseSAPValueToDouble(QCTotalH.Value) - HelperMethods.ParseValueToDouble(previousQty) + HelperMethods.ParseValueToDouble(opNewQty.Value));

                }
                if (opResc.ToString().StartsWith("SUBCON") && previousResc.ToString().StartsWith("SUBCON"))
                {
                    HelperMethods.UpdateEditText(uIAPIRawForm, "QCTotalSCF", HelperMethods.ParseSAPValueToDouble(QCTotalSCF.Value) - HelperMethods.ParseValueToDouble(previousLineTotal) + HelperMethods.ParseValueToDouble(newCost));
                    HelperMethods.UpdateEditText(uIAPIRawForm, "QCTotalSC", HelperMethods.ParseSAPValueToDouble(QCTotalSC.Value) - HelperMethods.ParseValueToDouble(previousLineTotal) + HelperMethods.ParseValueToDouble(newCost));
                }
                if (opResc.ToString().StartsWith("SUBCON") && !previousResc.ToString().StartsWith("SUBCON"))
                {
                    HelperMethods.UpdateEditText(uIAPIRawForm, "QCOPTot", HelperMethods.ParseSAPValueToDouble(QCOPTot.Value) - HelperMethods.ParseValueToDouble(previousLineTotal));
                    HelperMethods.UpdateEditText(uIAPIRawForm, "QCOpA", HelperMethods.ParseSAPValueToDouble(QCOpA.Value) - HelperMethods.ParseValueToDouble(previousLineTotal));
                    //HelperMethods.UpdateEditText(uIAPIRawForm, "QCOTCost", HelperMethods.ParseSAPValueToDouble(QCOTCost.Value) - HelperMethods.ParseValueToDouble(previousLineTotal) + HelperMethods.ParseValueToDouble(newCost));

                    HelperMethods.UpdateEditText(uIAPIRawForm, "QCTotalHF", HelperMethods.ParseSAPValueToDouble(QCTotalHF.Value) - HelperMethods.ParseValueToDouble(previousQty));
                    HelperMethods.UpdateEditText(uIAPIRawForm, "QCTotalH", HelperMethods.ParseSAPValueToDouble(QCTotalH.Value) - HelperMethods.ParseValueToDouble(previousQty));

                    HelperMethods.UpdateEditText(uIAPIRawForm, "QCTotalSCF", HelperMethods.ParseSAPValueToDouble(QCTotalSCF.Value) + HelperMethods.ParseValueToDouble(newCost));
                    HelperMethods.UpdateEditText(uIAPIRawForm, "QCTotalSC", HelperMethods.ParseSAPValueToDouble(QCTotalSC.Value) + HelperMethods.ParseValueToDouble(newCost));

                }
            }
            HelperMethods.UpdateEditText(uIAPIRawForm, "QCTEst", HelperMethods.ParseSAPValueToDouble(QCTEst.Value) - HelperMethods.ParseValueToDouble(previousLineTotal) + HelperMethods.ParseValueToDouble(newCost));
        }

        internal static void OperationsCalcTotal(IForm uIAPIRawForm)
        {
            double totalop = 0;
            double totalsub = 0;
            double totalqty = 0;
            double totalOC = 0;

            var mOperations = HelperMethods.GetMatrix(uIAPIRawForm, "mOper");

            // Ensure the datasource is updated with the UI values
            mOperations.FlushToDataSource();

            SAPbouiCOM.DBDataSource ds = (SAPbouiCOM.DBDataSource)uIAPIRawForm.DataSources.DBDataSources.Item("@STXQC19O");

            XElement xml = XElement.Parse(ds.GetAsXML());

            var rows = xml.Element("rows").Elements("row");

            foreach (var row in rows)
            {
                var cells = row.Element("cells").Elements("cell");

                string opRescValue = cells.FirstOrDefault(c => c.Element("uid").Value == "U_resCode")?.Element("value")?.Value;

                if (!string.IsNullOrEmpty(opRescValue) && !opRescValue.StartsWith("SUBCON"))
                {
                    string optotalValue = cells.FirstOrDefault(c => c.Element("uid").Value == "U_LineTot")?.Element("value")?.Value;
                    totalop += HelperMethods.ParseValueToDouble(optotalValue);

                    string qtytotalValue = cells.FirstOrDefault(c => c.Element("uid").Value == "U_Quantity")?.Element("value")?.Value;
                    totalqty += HelperMethods.ParseValueToDouble(qtytotalValue);
                }
                else
                {
                    string subtotalValue = cells.FirstOrDefault(c => c.Element("uid").Value == "U_LineTot")?.Element("value")?.Value;
                    totalsub += HelperMethods.ParseValueToDouble(subtotalValue);
                }
            }


            HelperMethods.UpdateEditText(uIAPIRawForm, "QCOPTot", totalop);
            HelperMethods.UpdateEditText(uIAPIRawForm, "QCOpA", totalop);
            HelperMethods.UpdateEditText(uIAPIRawForm, "QCTotalSCF", totalsub);
            HelperMethods.UpdateEditText(uIAPIRawForm, "QCTotalSC", totalsub);
            HelperMethods.UpdateEditText(uIAPIRawForm, "QCOTCost", totalOC);
            HelperMethods.UpdateEditText(uIAPIRawForm, "QCTotalH", totalqty);
            HelperMethods.UpdateEditText(uIAPIRawForm, "QCTotalHF", totalqty);

            HelperMethods.UpdateEditText(uIAPIRawForm, "QCTEst", totalop + totalOC + totalsub);

        }
    }
}