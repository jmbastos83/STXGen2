using SAPbouiCOM;

namespace STXGen2
{
    internal class MatrixOperations
    {
        public static void BindDataToMatrix(IForm uIAPIRawForm)
        {
            uIAPIRawForm.Freeze(true);
            SAPbouiCOM.Matrix mOperations = (SAPbouiCOM.Matrix)uIAPIRawForm.Items.Item("mOper").Specific;

            mOperations.Clear();

            mOperations.Columns.Item("#").DataBind.SetBound(true, "@STXQC19O", "VisOrder");
            mOperations.Columns.Item("OPTexture").DataBind.SetBound(true, "@STXQC19O", "U_Texture");
            mOperations.Columns.Item("OPResc").DataBind.SetBound(true, "@STXQC19O", "U_resCode");
            mOperations.Columns.Item("OPResN").DataBind.SetBound(true, "@STXQC19O", "U_resName");
            mOperations.Columns.Item("OPcode").DataBind.SetBound(true, "@STXQC19O", "U_opCode");
            mOperations.Columns.Item("OPName").DataBind.SetBound(true, "@STXQC19O", "U_opDesc");
            mOperations.Columns.Item("OPNameL").DataBind.SetBound(true, "@STXQC19O", "U_opDescL");
            mOperations.Columns.Item("OPStdT").DataBind.SetBound(true, "@STXQC19O", "U_sugQty");
            mOperations.Columns.Item("OPQtdT").DataBind.SetBound(true, "@STXQC19O", "U_Quantity");
            mOperations.Columns.Item("OPUom").DataBind.SetBound(true, "@STXQC19O", "U_UOM");
            mOperations.Columns.Item("OPCost").DataBind.SetBound(true, "@STXQC19O", "U_Price");
            mOperations.Columns.Item("OPTotal").DataBind.SetBound(true, "@STXQC19O", "U_LineTot");
            mOperations.Columns.Item("OPErrMsg").DataBind.SetBound(true, "@STXQC19O", "U_ErrMsg");

            mOperations.LoadFromDataSource();

            uIAPIRawForm.Freeze(false);
        }
    }
}