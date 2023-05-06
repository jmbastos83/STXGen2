using System;
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

namespace STXGen2
{
    internal class QCEvents
    {

        public readonly object MatrixLock = new object();
        public static SAPbouiCOM.DataTable operations;
        public static bool _processChooseFromList = false;
        public static Dictionary<int, string> _pendingCFLUpdates = new Dictionary<int, string>();

        public static string defValue { get; set; }
        public static object QCLength { get; private set; }
        public static int processOperationsListErr { get; private set; } = 0;
        public static bool operationsUpdate { get; set; } = false;


        public static void AddLineToOperationMatrix(SAPbouiCOM.Form oForm, Matrix operationsMatrix, int selectedMatrixRow)
        {
            throw new NotImplementedException();
        }

        public static void AddLineToTexturesMatrix(SAPbouiCOM.Form oForm, SAPbouiCOM.Matrix texturesMatrix, int selectedRow)
        {
            oForm.Freeze(true);

            if (texturesMatrix.RowCount == 5)
            {
                Program.SBO_Application.SetStatusBarMessage("Maximum number of textures reached.",BoMessageTime.bmt_Medium,false);
                return;
            }
            if (texturesMatrix.RowCount == 0 || selectedRow == texturesMatrix.RowCount)
            {
                texturesMatrix.AddRow();
                texturesMatrix.ClearRowData(selectedRow + 1);

                SAPbouiCOM.DBDataSource oDBDataSource = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("@STXQC19T");

                oDBDataSource.InsertRecord(selectedRow);
                oDBDataSource.SetValue("VisOrder", selectedRow, (selectedRow + 1).ToString());

                texturesMatrix.LoadFromDataSource();

            }
            else
            {
                if (selectedRow == -1 || selectedRow > texturesMatrix.RowCount)
                {
                    selectedRow = texturesMatrix.RowCount;
                }

                //Console.WriteLine($"Selected row: {selectedRow}");
                //Console.WriteLine($"Matrix row count: {texturesMatrix.RowCount}");

                if (texturesMatrix.RowCount < 5)
                {
                    // Store the values of the rows to be moved down
                    List<Dictionary<int, string>> rowsData = new List<Dictionary<int, string>>();
                    for (int rowIndex = selectedRow + 1; rowIndex <= texturesMatrix.RowCount; rowIndex++)
                    {
                        //Console.WriteLine($"Storing row: {rowIndex}");
                        rowsData.Add(new Dictionary<int, string>());
                        for (int colIndex = 0; colIndex < texturesMatrix.Columns.Count; colIndex++)
                        {
                            var cell = texturesMatrix.Columns.Item(colIndex).Cells.Item(rowIndex).Specific;
                            if (cell is SAPbouiCOM.EditText editText)
                            {
                                if (texturesMatrix.Columns.Item(colIndex).UniqueID == "#")
                                {
                                    rowsData[rowIndex - (selectedRow + 1)][colIndex] = (int.Parse(editText.Value) + 1).ToString();
                                }
                                else
                                {
                                    rowsData[rowIndex - (selectedRow + 1)][colIndex] = editText.Value.ToString();
                                }

                            }
                            if (cell is SAPbouiCOM.LinkedButton linkButton)
                            {
                                rowsData[rowIndex - (selectedRow + 1)][colIndex] = ((SAPbouiCOM.EditText)linkButton.Item.Specific).Value;
                            }
                            if (cell is SAPbouiCOM.ComboBox comboBox)
                            {
                                rowsData[rowIndex - (selectedRow + 1)][colIndex] = comboBox.Selected.Value.ToString();
                            }
                        }
                    }

                    // Add a new row in the data source
                    SAPbouiCOM.DBDataSource oDBDataSource = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("@STXQC19T");

                    oDBDataSource.InsertRecord(selectedRow + 1);
                    oDBDataSource.SetValue("VisOrder", selectedRow + 1, (selectedRow + 1).ToString());


                    // Set the LineID value to the max LineID + 1
                    int maxLineID = oDBDataSource.Size == 0 ? 0 : (int)QuoteCalculator.mtxMaxLineID;
                    oDBDataSource.SetValue("LineID", selectedRow + 1, (maxLineID + 1).ToString());
                    QuoteCalculator.mtxMaxLineID = maxLineID + 1;

                    // Refresh the matrix
                    texturesMatrix.LoadFromDataSource();

                    texturesMatrix.ClearRowData(selectedRow + 1);

                    texturesMatrix.FlushToDataSource();


                    
                    for (int rowIndex = 0; rowIndex < rowsData.Count; rowIndex++)
                    {

                        //Console.WriteLine($"Restoring row: {rowIndex}");
                        for (int colIndex = 0; colIndex < texturesMatrix.Columns.Count; colIndex++)
                        {
                            if (texturesMatrix.Columns.Item(colIndex).Type == BoFormItemTypes.it_EDIT)
                            {
                                ((SAPbouiCOM.EditText)texturesMatrix.Columns.Item(colIndex).Cells.Item(rowIndex + (selectedRow + 2)).Specific).Value = rowsData[rowIndex].ContainsKey(colIndex) ? rowsData[rowIndex][colIndex] : "";
                                if (colIndex == 0)
                                {
                                    texturesMatrix.FlushToDataSource();
                                }
                            }
                            if (texturesMatrix.Columns.Item(colIndex).Type == BoFormItemTypes.it_COMBO_BOX)
                            {
                                ((SAPbouiCOM.ComboBox)texturesMatrix.Columns.Item(colIndex).Cells.Item(rowIndex + (selectedRow + 2)).Specific).Select(rowsData[rowIndex].ContainsKey(colIndex) ? rowsData[rowIndex][colIndex] : "0", BoSearchKey.psk_ByValue);
                            }
                            if (texturesMatrix.Columns.Item(colIndex).Type == BoFormItemTypes.it_LINKED_BUTTON)
                            {
                                oDBDataSource.SetValue("U_Texture", (rowIndex + (selectedRow + 1)), rowsData[rowIndex].ContainsKey(colIndex) ? rowsData[rowIndex][colIndex] : "");
                                texturesMatrix.LoadFromDataSource();
                            }
                        }

                    }
                    texturesMatrix.SetCellFocus(selectedRow, 1);
                }
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
            SAPbouiCOM.Form oForm = ((SAPbouiCOM.Form)uIAPIRawForm);
            SAPbouiCOM.Matrix mTextures = (SAPbouiCOM.Matrix)oForm.Items.Item("mTextures").Specific;

            // Loop through each row in the mtxQCItems Matrix control
            for (int i = 1; i <= mTextures.RowCount; i++)
            {

                if (isUoMAreaChanging == true)
                {
                    SAPbouiCOM.EditText CovArea = (SAPbouiCOM.EditText)mTextures.Columns.Item("QCCovA").Cells.Item(i).Specific;
                    SAPbouiCOM.EditText calArea = (SAPbouiCOM.EditText)oForm.Items.Item("QCArea").Specific;

                    CovArea.Value = calArea.Value;
                }
                else
                {
                    SAPbouiCOM.EditText CovArea = (SAPbouiCOM.EditText)mTextures.Columns.Item("QCCovA").Cells.Item(i).Specific;
                    string treatedCovArea = (string)Regex.Replace(CovArea.Value, @"(\w)2", match => $"{match.Groups[1]}²");

                    double covA = double.Parse(Regex.Replace((string.IsNullOrEmpty(treatedCovArea) ? "0" : treatedCovArea), $@"[^\d{Utils.decSep}{Utils.thousSep}]", ""));
                    double convertedcovA = DBCalls.ConvertAreaDimensions(covA, selectedUOM, previousUOM);
                    CovArea.Value = $"{convertedcovA} {selectedUOM}²";
                }

            }
        }

        internal static string CalculateArea(string formUID, string selectedUoM)
        {
            SAPbouiCOM.Form oForm = Program.SBO_Application.Forms.Item(formUID);
            EditText edtQCLength = (EditText)oForm.Items.Item("QCLength").Specific;
            EditText edtQCWidth = (EditText)oForm.Items.Item("QCWidth").Specific;
            EditText edtQCArea = (EditText)oForm.Items.Item("QCArea").Specific;

            float length = float.Parse(Regex.Replace((string.IsNullOrEmpty(edtQCLength.Value) ? "0" : edtQCLength.Value), $@"[^\d{Utils.decSep}{Utils.thousSep}]", ""));
            float width = float.Parse(Regex.Replace((string.IsNullOrEmpty(edtQCWidth.Value) ? "0" : edtQCWidth.Value), $@"[^\d{Utils.decSep}{Utils.thousSep}]", ""));


            float area = length * width;
            edtQCArea.Value = (area.ToString() + " " + selectedUoM + "²");

            return edtQCArea.Value;
        }

        internal static string LoadImageFromResources()
        {
            string imagePath = "";
            Assembly assembly = Assembly.GetExecutingAssembly();
            string resourceName = "STXGen2.Properties.Light-green.jpg";

            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream != null)
                {
                    Image image = Image.FromStream(stream);
                    imagePath = Path.GetTempPath() + "Light-green.jpg";
                    image.Save(imagePath);
                }
            }

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
            if (QuoteCalculator.selectedMatrixRow > 0)
            {
                SAPbouiCOM.DBDataSource oDBDataSource = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("@STXQC19O");

                // Remove the row from the data source
                oDBDataSource.RemoveRecord(QuoteCalculator.selectedMatrixRow - 1);

                // Update the # aka LineID column
                for (int i = QuoteCalculator.selectedMatrixRow - 1; i < oDBDataSource.Size; i++)
                {
                    oDBDataSource.SetValue("VisOrder", i, (i + 1).ToString());
                }

                // Refresh the matrix
                OperationsMatrix.LoadFromDataSource();

                // Set this property to prevent the selection of the row after deletion.
                OperationsMatrix.SelectionMode = BoMatrixSelect.ms_None;

                // Set the selection mode back to the default after loading the data source.
                OperationsMatrix.SelectionMode = BoMatrixSelect.ms_Auto;
            }
        }

        internal static void RemoveLinefromTexturesMatrix(SAPbouiCOM.Form oForm, Matrix texturesMatrix, int selectedMatrixRow)
        {
            if (QuoteCalculator.selectedMatrixRow > 0)
            {
                SAPbouiCOM.DBDataSource oDBDataSource = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("@STXQC19T");

                // Remove the row from the data source
                oDBDataSource.RemoveRecord(QuoteCalculator.selectedMatrixRow - 1);

                // Update the # aka LineID column
                for (int i = QuoteCalculator.selectedMatrixRow - 1; i < oDBDataSource.Size; i++)
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

        internal static (string AdditionalConditions, string ConcatenatedTextureCodes, string tClassExpression, string OpQuantityExpression) GetAdditionalConditions(List<Dictionary<string, string>> matrix1Values)
        {
            string quantity = "";
            string textureCode = "";
            string coverageArea = "0";
            string tClass = "";
            string GeoComplex = "";

            var calcFactorList = new List<string>();
            var TextureClassList = new List<string>();
            var OpQuantityList = new List<string>();
            string concatenatedTextureCodes = GetConcatenatedTextureCodes(matrix1Values);

            for (int i = 0; i < matrix1Values.Count; i++)
            {
                Dictionary<string, string> rowValues = matrix1Values[i];

                textureCode = rowValues["QCTexture"];
                quantity = rowValues["QCQuantity"];
                tClass = rowValues["QCTClass"];
                GeoComplex = rowValues["QCGComp"];
                coverageArea = Regex.Replace(rowValues["QCCovA"], $@"[^\d{Utils.decSep}{Utils.thousSep}]", "");
                coverageArea = DBCalls.ConvertDimMeters(double.Parse(coverageArea), QuoteCalculator.selectedUOM);

                string condition1 = $"WHEN T2.\"U_standexReference\" = '{textureCode}' AND T1.\"U_STXQtyBy\" = 'A' THEN {coverageArea}";
                calcFactorList.Add(condition1);

                string condition2 = $"WHEN R0.\"Texture\" = '{textureCode}' THEN (select \"U_Factor\" from \"@STXQC19TCLASS\" Where \"Code\" = {tClass})";
                TextureClassList.Add(condition2);

                string condition3 = $"WHEN {GeoComplex} = 1 then T1.\"Quantity\" when {GeoComplex} = 2 then T1.\"U_STXQTYGC2\" when {GeoComplex} = 3 then T1.\"U_STXQTYGC3\"";
                OpQuantityList.Add(condition3);



            }

            string calcFactorcond = string.Join(" ", calcFactorList);
            string calcFactorExpression = $"(CASE {calcFactorcond} ELSE {quantity} END) as \"CalcFactor\"";

            string tClasscond = string.Join(" ", TextureClassList);
            string tClassExpression = $"(CASE {tClasscond} ELSE 1 END) as \"TClassFactor\"";

            string OPQty = string.Join(" ", OpQuantityList);
            string OpQuantityExpression = $"(CASE {OPQty} ELSE 0 END) as \"Quantity\"";

            return (calcFactorExpression, concatenatedTextureCodes, tClassExpression, OpQuantityExpression);
        }


        private static string GetConcatenatedTextureCodes(List<Dictionary<string, string>> matrix1Values)
        {
            var textureCodes = matrix1Values.Select(x => $"'{x["QCTexture"]}'");
            return string.Join(",", textureCodes);
        }

        internal static void GetOperations(IForm uIAPIRawForm)
        {
            processOperationsListErr = 0;

            SAPbouiCOM.Matrix mOperations = (Matrix)uIAPIRawForm.Items.Item("mOper").Specific;
            SAPbouiCOM.Matrix matrix1 = (Matrix)uIAPIRawForm.Items.Item("mTextures").Specific;

            List<Dictionary<string, string>> matrix1Values = QCEvents.GetAllValuesFromMatrix1(matrix1);

            processOperationsList(uIAPIRawForm, matrix1Values);

            switch (processOperationsListErr)
            {
                case 0:
                    processMTOperationsList(uIAPIRawForm, mOperations, matrix1Values);
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

        private static void processMTOperationsList(IForm uIAPIRawForm, Matrix mOperations, List<Dictionary<string, string>> mtTexture)
        {


            bool dataTableExists = false;

            var (CalcFactorConditions, concatenatedTextureCodes, tclassConditions, OpQuantityExpression) = QCEvents.GetAdditionalConditions(mtTexture);

            // Create a unique identifier for the DataTable
            string dataTableID = "Operations";


            // Check if the DataTable with the ID "Operations" exists
            for (int i = 0; i < uIAPIRawForm.DataSources.DataTables.Count; i++)
            {
                SAPbouiCOM.DataTable dt = uIAPIRawForm.DataSources.DataTables.Item(i);
                if (dt.UniqueID == dataTableID)
                {
                    dataTableExists = true;
                    break;
                }
            }

            if (!dataTableExists)
            {
                uIAPIRawForm.DataSources.DataTables.Add(dataTableID);
                operations = uIAPIRawForm.DataSources.DataTables.Item(dataTableID);
            }
            else
            {
                operations = uIAPIRawForm.DataSources.DataTables.Item(dataTableID);
            }

            Dictionary<string, int> mOperationsRowIndexMap = new Dictionary<string, int>();
            for (int j = 1; j <= mOperations.RowCount; j++)
            {
                string existingVisOrder = ((SAPbouiCOM.EditText)mOperations.Columns.Item("#").Cells.Item(j).Specific).Value;
                mOperationsRowIndexMap[existingVisOrder] = j;
            }

            DBCalls.GetOperation(operations, uIAPIRawForm, mOperations, CalcFactorConditions, concatenatedTextureCodes, tclassConditions, OpQuantityExpression, ((SAPbouiCOM.EditText)uIAPIRawForm.Items.Item("QCSubPart").Specific).Value);

            int operationscount = operations.Rows.Count;
            if (operationscount > 0)
            {
                try
                {
                    uIAPIRawForm.Freeze(true);

                    mOperations.Clear();

                    // Bind the DataTable columns to the matrix columns
                    mOperations.Columns.Item("#").DataBind.Bind(dataTableID, "VisOrder");
                    mOperations.Columns.Item("OPTexture").DataBind.Bind(dataTableID, "OPTexture");
                    mOperations.Columns.Item("OPResc").DataBind.Bind(dataTableID, "OPResc");
                    mOperations.Columns.Item("OPResN").DataBind.Bind(dataTableID, "OPResN");
                    mOperations.Columns.Item("OPcode").DataBind.Bind(dataTableID, "OPcode");
                    mOperations.Columns.Item("OPName").DataBind.Bind(dataTableID, "OPName");
                    mOperations.Columns.Item("OPNameL").DataBind.Bind(dataTableID, "OPNameL");
                    mOperations.Columns.Item("OPStdT").DataBind.Bind(dataTableID, "OPStdT");
                    mOperations.Columns.Item("OPQtdT").DataBind.Bind(dataTableID, "OPQtdT");
                    mOperations.Columns.Item("OPUom").DataBind.Bind(dataTableID, "OPUom");
                    mOperations.Columns.Item("OPCost").DataBind.Bind(dataTableID, "OPCost");
                    mOperations.Columns.Item("OPTotal").DataBind.Bind(dataTableID, "OPTotal");
                    mOperations.Columns.Item("OPErrMsg").DataBind.Bind(dataTableID, "OPErrMsg");

                    // Bind check boxes using UserDataSources
                    for (int i = 0; i < operationscount; i++)
                    {
                        string checkUDSId = $"CheckUDS{i}";
                        if (!uIAPIRawForm.DataSources.UserDataSources.Cast<SAPbouiCOM.UserDataSource>().Any(uds => uds.UID == checkUDSId))
                        {
                            uIAPIRawForm.DataSources.UserDataSources.Add(checkUDSId, BoDataType.dt_SHORT_TEXT, 1);
                        }
                        mOperations.Columns.Item("OPcheck").DataBind.SetBound(true, "", checkUDSId);

                    }

                    // Load data from the DataTable to the matrix
                    mOperations.LoadFromDataSource();

                    if (operations.Rows.Count < mOperationsRowIndexMap.Count)
                    {
                        for (int j = mOperationsRowIndexMap.Count; j > operations.Rows.Count; j--)
                        {
                            mOperations.DeleteRow(j);
                        }
                        mOperations.FlushToDataSource();
                    }

                    mOperations.AutoResizeColumns();
                    SetMatrixRowColor(mOperations, "OPErrMsg");
                    uIAPIRawForm.Freeze(false);

                }
                catch (Exception ex)
                {
                    Program.SBO_Application.SetStatusBarMessage(ex.Message, BoMessageTime.bmt_Medium, true);
                }
                finally
                {
                    operationsUpdate = true;

                    mOperations.FlushToDataSource();
                    
                    uIAPIRawForm.Mode = BoFormMode.fm_UPDATE_MODE;

                    Program.SBO_Application.SetStatusBarMessage("Operations matrix updated.", BoMessageTime.bmt_Short, false);
                }
            }
        }

        private static void SetMatrixRowColor(SAPbouiCOM.Matrix mOperations, string colUID)
        {
            Color orangeColor = Color.FromArgb(0xFF, 0xD1, 0x55);
            int warning = (orangeColor.R) + (orangeColor.G << 8) + (orangeColor.B << 16);

            for (int rowIndex = 1; rowIndex <= mOperations.RowCount; rowIndex++)
            {
                SAPbouiCOM.EditText cell = (SAPbouiCOM.EditText)mOperations.Columns.Item(colUID).Cells.Item(rowIndex).Specific;
                if (!string.IsNullOrEmpty(cell.Value))
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
                uIAPIRawForm.Freeze(true);

                SAPbouiCOM.DataTable operations;

                SAPbouiCOM.Matrix mOperations = (Matrix)uIAPIRawForm.Items.Item("mOper").Specific;
                mOperations.Clear();

                SAPbouiCOM.Matrix matrix1 = (Matrix)uIAPIRawForm.Items.Item("mTextures").Specific;
                List<Dictionary<string, string>> matrix1Values = QCEvents.GetAllValuesFromMatrix1(matrix1);


                var (CalcFactorCond, concatenatedTextureCodes, tClassCond, OpQuantityExpression) = QCEvents.GetAdditionalConditions(matrix1Values);

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
            finally
            {
                uIAPIRawForm.Freeze(false);
            }

        }
    }
}