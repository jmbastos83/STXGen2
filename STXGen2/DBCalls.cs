using SAPbobsCOM;
using SAPbouiCOM;
using STXGen2.Properties;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Xml.Linq;

namespace STXGen2
{
    internal class DBCalls
    {

        public string Quantity { get; set; }
        public string GeoComplex { get; set; }



        internal static void GetFilterOperations(SAPbouiCOM.ComboBox comboBox, SAPbouiCOM.EditText qCDocEntry)
        {
            string query = "select - 1 as \"Code\", 'All Tasks' as \"Description\"\n" +
                          "union all\n" +
                          "select distinct \"U_seq\" as \"Code\", Case when \"U_seq\" = 0 then 'Initial Tasks' when U_seq = 99 then 'Final Tasks' else concat('Texture: ', \"U_Texture\") end as \"Description\" from \"@STXQC19O\" where \"DocEntry\" = {0}";
            query = string.Format(query, qCDocEntry.Value);
            Recordset rs = Utils.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset) as Recordset;
            rs.DoQuery(query);

            while (!rs.EoF)
            {
                string value = rs.Fields.Item("Code").Value.ToString();
                string descr = rs.Fields.Item("Description").Value.ToString();
                comboBox.ValidValues.Add(value, descr);
                rs.MoveNext();
            }
        }



        internal static string GetUserLanguage()
        {
            string sSql = $"SELECT \"Language\" FROM OUSR WHERE \"USER_CODE\" = '" + Utils.oCompany.UserName + "'";

            Recordset rs = Utils.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset) as Recordset;
            rs.DoQuery(sSql);

            string langCode = rs.Fields.Item("Language").Value.ToString();

            return (langCode);
        }

        internal static (string qty, string TextClass) GetTextureInfo(string selectedItemCode)
        {
            string qty = "";
            string textClass = "";

            string sSql = $"select \"Code\",1 as \"Quantity\", \"U_complexityIX\" from \"@STXSETPTEXTURES\" where \"Code\" = '{selectedItemCode}'";

            Recordset rs = Utils.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset) as Recordset;
            rs.DoQuery(sSql);

            if (!rs.EoF)
            {

                qty = rs.Fields.Item("Quantity").Value.ToString();
                textClass = rs.Fields.Item("U_complexityIX").Value.ToString();

                return (qty, textClass);
            }
            else
            {
                return ("1", "");
            }
        }

        internal static double ConvertDimensions(double size, string selectedUoM, string previousUom)
        {
            double oldFactor = 0;
            double newFactor = 0;
            double result = 0;

            string sSql = $"select \"UnitDisply\",\"SizeInMM\" from \"OLGT\" where \"UnitDisply\" = '{selectedUoM.Replace("'", "''")}'";
            Recordset rs = Utils.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset) as Recordset;
            rs.DoQuery(sSql);
            if (!rs.EoF)
            {
                newFactor = (double)rs.Fields.Item("SizeInMM").Value;
            }

            string sSql2 = $"select \"UnitDisply\",\"SizeInMM\" from \"OLGT\" where \"UnitDisply\" = '{previousUom.Replace("'", "''")}'";
            Recordset rs2 = Utils.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset) as Recordset;
            rs2.DoQuery(sSql2);

            if (!rs2.EoF)
            {
                oldFactor = (double)rs2.Fields.Item("SizeInMM").Value;
            }

            result = (size * oldFactor) / newFactor;

            return result;
        }

        internal static int GetMatrixLastLineID(string qCDocEntry)
        {
            int maxLineID = 0;

            string sSql = $"select max(\"LineID\") as \"LineId\" from \"@STXQC19T\" where \"DocEntry\" = '{qCDocEntry}'";
            Recordset rs = Utils.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset) as Recordset;
            rs.DoQuery(sSql);
            if (!rs.EoF)
            {
                maxLineID = (int)rs.Fields.Item("LineId").Value;
            }

            return maxLineID;
        }

        internal static int GetMatrixOPLastLineID(string qCDocEntry)
        {
            int maxLineID = 0;

            string sSql = $"select max(\"LineID\") as \"LineId\" from \"@STXQC19O\" where \"DocEntry\" = '{qCDocEntry}'";
            Recordset rs = Utils.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset) as Recordset;
            rs.DoQuery(sSql);
            if (!rs.EoF)
            {
                maxLineID = (int)rs.Fields.Item("LineId").Value;
            }

            return maxLineID;
        }

        internal static void GetOperation(SAPbouiCOM.DataTable operations, IForm uIAPIRawForm, Matrix mOperations, string CalcFactor, string concatenatedTextureCodes, string tclassFactor, string OpQuantityExpression, string SptCode, bool DefBOM)
        {


            string query = "Select  ROW_NUMBER() OVER (ORDER BY X0.\"Order\",X0.\"Texture\",X0.\"U_groupOrder\",X0.\"U_operationOrder\") AS \"#\",X0.\"Texture\" as \"OPTexture\",X0.\"U_operationResource\" as \"OPResc\",X0.\"ResName\" as \"OPResN\",X0.\"U_operationCode\" as \"OPcode\",\n" +
                           "X0.\"U_STXOPDes\" as \"OPName\",X0.\"U_STXOPDesLocal\" as \"OPNameL\",cast(Round((Case when X0.\"U_STXQtyBy\" = 'A' then (X0.\"CalcFactor\" / X0.\"PlAvgSize\") * (X0.\"Quantity\" / X0.\"NTimes\") * X0.\"TClassFactor\" else (X0.\"Quantity\" / X0.\"NTimes\") * X0.\"TClassFactor\" end),{5}) AS DECIMAL(18, {5})) as \"OPStdT\",\n" +
                           "cast(Round((Case when X0.\"U_STXQtyBy\" = 'A' then (X0.\"CalcFactor\" / X0.\"PlAvgSize\") * (X0.\"Quantity\" / X0.\"NTimes\") * X0.\"TClassFactor\" else (X0.\"Quantity\" / X0.\"NTimes\") * X0.\"TClassFactor\" end),{5}) AS DECIMAL(18, {5})) as \"OPQtdT\",X0.\"UnitOfMsr\" as \"OPUom\",cast(Round((X0.\"ResCost\"),{6}) AS DECIMAL(18, {6})) as \"OPCost\",\n" +
                           "cast(Round(((Case when X0.\"U_STXQtyBy\" = 'A' then (X0.\"CalcFactor\" / X0.\"PlAvgSize\") * (X0.\"Quantity\" / X0.\"NTimes\") * X0.\"TClassFactor\" else (X0.\"Quantity\" / X0.\"NTimes\") * X0.\"TClassFactor\" end) * X0.\"ResCost\"),{6}) AS DECIMAL(18, {6})) as \"OPTotal\",\n" +
                           "Case when coalesce(isnull(X0.\"ResName\",''),'') = '' then '{8}' when coalesce(isnull(X0.\"U_STXOPDes\",''),'') = '' then '{9}' when(X0.\"U_STXQtyBy\" = 'A' and X0.\"CalcFactor\" = 0) then '{10}' when X0.\"ResCost\" = 0 then '{11}' end as \"OPErrMsg\",\n" +
                           "Case when X0.\"U_PlanType\" = 'I' then 0 when  X0.\"U_PlanType\" = 'F' then 99 else DENSE_RANK() OVER (order by X0.\"Texture\")-1 end as \"OPSeq\",'' as \"Dummy\" from(\n" +
                           "Select R0.\"Order\",CASE WHEN R0.\"U_PlanType\" = 'N' then R0.\"U_groupOrder\" else NULL END as \"U_groupOrder\",R0.\"U_operationOrder\", R0.\"U_PlanType\",R0.\"Texture\",R0.\"U_operationResource\",R1.\"ResName\",R0.\"U_operationCode\",\n" +
                           "R0.\"U_STXOPDes\",R0.\"U_STXOPDesLocal\",R0.\"PlAvgSize\",sum(R0.\"Quantity\") as \"Quantity\",R0.\"U_STXQtyBy\",R0.\"CalcFactor\",{3},R1.\"ResCost\",R1.\"UnitOfMsr\",sum(R0.\"NTimes\") as \"NTimes\"\n" +
                           "from(\n" +
                           "select  1 as \"Order\", T2.\"U_groupOrder\", Case When '{12}' = 'True' then T1.\"VisOrder\" else T2.\"U_operationOrder\" end as \"U_operationOrder\" , T3.\"U_PlanType\", Case when T3.\"U_PlanType\" = 'I' or T3.\"U_PlanType\" = 'F' then null else\n" +
                           "T2.\"U_standexReference\" end as \"Texture\",T2.\"U_operationResource\", T2.\"U_operationCode\", T3.\"U_STXOPDes\", T3.\"U_STXOPDesLocal\", {0},\n" +
                           "T1.\"U_STXQtyBy\",T0.\"PlAvgSize\",{4},1 as \"NTimes\"\n" +
                           "from OITT T0\n" +
                           "inner join ITT1 T1 on T0.\"Code\" = T1.\"Father\"\n" +
                           "inner join(select * from \"@STXSETPTEXTURETASKS\" where \"U_standexReference\" in ({1})) T2 on T1.\"U_STXOPCode\" = T2.\"U_operationCode\"\n" +
                           "left join \"@STXOPERATIONS\" T3 on T2.\"U_operationCode\"= T3.\"Code\"\n" +
                           "where T0.\"Code\" = '{2}' and T3.\"U_PlanType\" = 'I' \n" +

                           "union all\n" +

                           "select  2 as \"Order\", T2.\"U_groupOrder\",Case When '{12}' = 'True' then T1.\"VisOrder\" else T2.\"U_operationOrder\" end as \"U_operationOrder\", T3.\"U_PlanType\", Case when T3.\"U_PlanType\" = 'I' or T3.\"U_PlanType\" = 'F' then null else\n" +
                           "T2.\"U_standexReference\" end as \"Texture\",T2.\"U_operationResource\", T2.\"U_operationCode\", T3.\"U_STXOPDes\", T3.\"U_STXOPDesLocal\", {0},\n" +
                           "T1.\"U_STXQtyBy\",T0.\"PlAvgSize\",{4},1 as \"NTimes\"\n" +
                           "from OITT T0\n" +
                           "inner join ITT1 T1 on T0.\"Code\" = T1.\"Father\"\n" +
                           "inner join(select * from \"@STXSETPTEXTURETASKS\" where \"U_standexReference\" in ({1})) T2 on T1.\"U_STXOPCode\" = T2.\"U_operationCode\"\n" +
                           "left join \"@STXOPERATIONS\" T3 on T2.\"U_operationCode\"= T3.\"Code\"\n" +
                           "where T0.\"Code\" = '{2}' and T3.\"U_PlanType\" not in ('I', 'F') \n" +

                           "union all\n" +

                           "select  3 as \"Order\", T2.\"U_groupOrder\", Case When '{12}' = 'True' then T1.\"VisOrder\" else T2.\"U_operationOrder\" end as \"U_operationOrder\", T3.\"U_PlanType\", Case when T3.\"U_PlanType\" = 'I' or T3.\"U_PlanType\" = 'F' then null else\n" +
                           "T2.\"U_standexReference\" end as \"Texture\",T2.\"U_operationResource\", T2.\"U_operationCode\", T3.\"U_STXOPDes\", T3.\"U_STXOPDesLocal\", {0},\n" +
                           "T1.\"U_STXQtyBy\",T0.\"PlAvgSize\",{4},1 as \"NTimes\"\n" +
                           "from OITT T0\n" +
                           "inner join ITT1 T1 on T0.\"Code\" = T1.\"Father\"\n" +
                           "inner join(select * from \"@STXSETPTEXTURETASKS\" where \"U_standexReference\" in ({1})) T2 on T1.\"U_STXOPCode\" = T2.\"U_operationCode\"\n" +
                           "left join \"@STXOPERATIONS\" T3 on T2.\"U_operationCode\"= T3.\"Code\"\n" +
                           "where T0.\"Code\" = '{2}' and T3.\"U_PlanType\" = 'F') as R0\n" +
                           "left join (select \"ResCode\",\"ResName\",\"UnitOfMsr\",(\"StdCost1\"+\"StdCost2\"+\"StdCost3\"+\"StdCost4\"+\"StdCost5\"+\"StdCost6\"+\"StdCost7\"+\"StdCost8\"+\"StdCost9\"+\"StdCost10\") as \"ResCost\" from ORSC) R1 on R0.\"U_operationResource\" = R1.\"ResCode\"\n" +
                           "group by R0.\"Order\",CASE WHEN R0.\"U_PlanType\" = 'N' then R0.\"U_groupOrder\" else NULL END,R0.\"U_operationOrder\", R0.\"U_PlanType\",R0.\"Texture\",R0.\"U_operationResource\",R1.\"ResName\",R0.\"U_operationCode\", R0.\"U_STXOPDes\",R0.\"U_STXOPDesLocal\",R0.\"PlAvgSize\",R0.\"U_STXQtyBy\",R0.\"CalcFactor\",R1.\"ResCost\",R1.\"UnitOfMsr\") X0, OADM X1\n" +
                           "order by X0.\"Order\",X0.\"Texture\",X0.\"U_groupOrder\",X0.\"U_operationOrder\"";

            query = string.Format(query, CalcFactor, concatenatedTextureCodes, SptCode, tclassFactor, OpQuantityExpression, Utils.QtyDec, Utils.PriceDec, Utils.SumDec, Resources.mOperErr1, Resources.mOperErr2, Resources.mOperErr3, Resources.mOperErr4, DefBOM);



            try
            {
                operations.ExecuteQuery(query);

            }
            catch (Exception ex)
            {

                Program.SBO_Application.SetStatusBarMessage(ex.Message);
            }
        }

        internal static string ConvertDimMeters(double size, string selectedUOM)
        {
            double oldFactor = 0;
            double targetFactor = 0;
            double area = 0;
            double result = 0;

            string MeterUOM = "m";

            string sSql = $"select \"UnitDisply\",\"SizeInMM\" from \"OLGT\" where \"UnitDisply\" = '{MeterUOM.Replace("'", "''")}'";
            Recordset rs = Utils.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset) as Recordset;
            rs.DoQuery(sSql);
            if (!rs.EoF)
            {
                targetFactor = (double)rs.Fields.Item("SizeInMM").Value;
            }

            string sSql2 = $"select \"UnitDisply\",\"SizeInMM\" from \"OLGT\" where \"UnitDisply\" = '{selectedUOM.Replace("'", "''")}'";
            Recordset rs2 = Utils.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset) as Recordset;
            rs2.DoQuery(sSql2);

            if (!rs2.EoF)
            {
                oldFactor = (double)rs2.Fields.Item("SizeInMM").Value;
            }

            // Convert targetFactor to square millimeters
            targetFactor = targetFactor * targetFactor;

            area = size * oldFactor * oldFactor;

            // Calculate the result in the target UoM
            result = (area / targetFactor);

            CultureInfo customCultureInfo = new CultureInfo("en-US");
            customCultureInfo.NumberFormat.NumberDecimalSeparator = ".";
            customCultureInfo.NumberFormat.NumberGroupSeparator = "";

            return result.ToString(customCultureInfo);
        }

        internal static double ConvertAreaDimensions(double covA, string selectedUOM, string previousUOM)
        {
            double oldFactor = 0;
            double targetFactor = 0;
            double area = 0;
            double result = 0;

            string sSql2 = $"select \"UnitDisply\",\"SizeInMM\" from \"OLGT\" where \"UnitDisply\" = '{previousUOM.Replace("'", "''")}'";
            Recordset rs2 = Utils.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset) as Recordset;
            rs2.DoQuery(sSql2);

            if (!rs2.EoF)
            {
                oldFactor = (double)rs2.Fields.Item("SizeInMM").Value;
            }

            // Get the conversion factor for the target UoM
            string sSql3 = $"select \"UnitDisply\",\"SizeInMM\" from \"OLGT\" where \"UnitDisply\" = '{selectedUOM.Replace("'", "''")}'";
            Recordset rs3 = Utils.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset) as Recordset;
            rs3.DoQuery(sSql3);
            if (!rs3.EoF)
            {
                targetFactor = (double)rs3.Fields.Item("SizeInMM").Value;
            }

            // Convert targetFactor to square millimeters
            targetFactor = targetFactor * targetFactor;

            area = covA * oldFactor * oldFactor;

            // Calculate the result in the target UoM
            result = (area / targetFactor);

            return result;
        }

        internal static void UpdateOperationsDB(System.Data.DataTable mOperations, string qCDocEntry)
        {
            QCEvents.operationsUpdate = false;

            SAPbobsCOM.GeneralData oChild = null;
            SAPbobsCOM.GeneralDataCollection oChildren = null;

            //SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)uIAPIRawForm.Items.Item("mOper").Specific;

            SAPbobsCOM.CompanyService oCompanyService = Utils.oCompany.GetCompanyService();
            SAPbobsCOM.GeneralService oGeneralService = oCompanyService.GetGeneralService("STXQC19");
            SAPbobsCOM.GeneralData oGeneralData = (SAPbobsCOM.GeneralData)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
            SAPbobsCOM.GeneralDataParams oGeneralParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);


            for (int i = 0; i < mOperations.Rows.Count; i++)
            {
                try
                {
                    oGeneralParams.SetProperty("DocEntry", qCDocEntry);      //Primary Key
                    oGeneralData = oGeneralService.GetByParams(oGeneralParams);

                    oChildren = oGeneralData.Child("STXQC19O");

                    // Check if the child at index i exists
                    if (i < oChildren.Count)
                    {
                        // If it exists, retrieve it
                        oChild = oChildren.Item(i);
                    }
                    else
                    {
                        // If it doesn't exist, add a new child and then retrieve it
                        oChildren.Add();
                        oChild = oChildren.Item(oChildren.Count - 1);
                    }

                    oChild.SetProperty("U_Texture", mOperations.Rows[i]["OPTexture"]);
                    oChild.SetProperty("U_resCode", mOperations.Rows[i]["OPResc"]);
                    oChild.SetProperty("U_resName", mOperations.Rows[i]["OPResN"]);
                    oChild.SetProperty("U_opCode", mOperations.Rows[i]["OPcode"]);
                    oChild.SetProperty("U_opDesc", mOperations.Rows[i]["OPName"]);
                    oChild.SetProperty("U_opDescL", mOperations.Rows[i]["OPNameL"]);
                    oChild.SetProperty("U_sugQty", mOperations.Rows[i]["OPStdT"]);
                    oChild.SetProperty("U_Quantity", mOperations.Rows[i]["OPQtdT"]);
                    oChild.SetProperty("U_UOM", mOperations.Rows[i]["OPUom"]);
                    oChild.SetProperty("U_Price", mOperations.Rows[i]["OPCost"]);
                    oChild.SetProperty("U_LineTot", mOperations.Rows[i]["OPTotal"]);
                    oChild.SetProperty("U_ErrMsg", mOperations.Rows[i]["OPErrMsg"]);
                    oChild.SetProperty("U_seq", mOperations.Rows[i]["OPSeq"]);

                    //Update the UDO Record                
                    oGeneralService.Update(oGeneralData);   // If Child Table does not have any record, it will create; else, update the existing one

                }
                catch (Exception ex)
                {
                    Program.SBO_Application.SetStatusBarMessage(ex.Message, BoMessageTime.bmt_Medium, true);
                }
            }
            for (int j = oChildren.Count - 1; j >= mOperations.Rows.Count; j--)
            {
                oChildren.Remove(j);
                oGeneralService.Update(oGeneralData);
            }
            //Program.SBO_Application.SetStatusBarMessage("Operations imported sucessfully.", BoMessageTime.bmt_Medium, false);
        }

        internal static (string,string) GetSPT(SAPbouiCOM.EditText qCSubPart)
        {
            string spt = "";
            string descr = "";

            string sSql = $"SELECT T0.\"ItemCode\", T0.\"ItemName\" as \"Part Name\" FROM OITM T0 WHERE T0.\"ItemCode\" ='{qCSubPart.Value}'";
            Recordset rs = Utils.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset) as Recordset;
            rs.DoQuery(sSql);

            if (!rs.EoF)
            {
                descr = (string)rs.Fields.Item("Part Name").Value;
            }

            QuoteCalculator.parttDescr = descr;
            descr = descr + ": ";

            string newQCSubPart = qCSubPart.String.Length > 2 ? qCSubPart.String.Substring(0, qCSubPart.String.Length - 2) + "00": qCSubPart.String + "00";
            string sSql2 = $"SELECT T0.\"ItemCode\", T0.\"ItemName\" as \"Part Name\" FROM OITM T0 WHERE T0.\"ItemCode\" like '{newQCSubPart}'";

            Recordset rs2 = Utils.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset) as Recordset;
            rs2.DoQuery(sSql2);

            if (!rs2.EoF)
            {
                spt = (string)rs2.Fields.Item("ItemCode").Value;
                descr = descr + (string)rs2.Fields.Item("Part Name").Value;
            }
            return (spt,descr);
        }

        internal static string ResCost(string resCode)
        {
            string cost = "";
            string sSql = $"select \"ResCode\",\"ResName\",\"UnitOfMsr\",(\"StdCost1\"+\"StdCost2\"+\"StdCost3\"+\"StdCost4\"+\"StdCost5\"+\"StdCost6\"+\"StdCost7\"+\"StdCost8\"+\"StdCost9\"+\"StdCost10\") as \"ResCost\" from ORSC WHERE \"ResCode\" ='{resCode}'";
            Recordset rs = Utils.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset) as Recordset;
            rs.DoQuery(sSql);

            if (!rs.EoF)
            {
                double resCost = (double)rs.Fields.Item("ResCost").Value;
                cost = resCost.ToString(CultureInfo.InvariantCulture);
            }
            return cost;
        }
    }
}