using SAPbobsCOM;
using SAPbouiCOM;
using STXGen2.Properties;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace STXGen2
{
    internal class DBCalls
    {

        public string Quantity { get; set; }
        public string GeoComplex { get; set; }

        internal static void DocumentTrackerInfo(SAPbouiCOM.DataTable dataTable,string openDocEntry)
        {
            string query = "SELECT ROW_NUMBER() OVER(ORDER BY T0.\"DocNum\",T1.\"LineNum\") AS \"#\", Case When coalesce(WO.\"DocEntry\",'') <> '' then 'N' else 'Y' end as \"Check\",\n" +
                            "T0.\"DocNum\" as \"SONum\",T1.\"LineNum\" as \"docLine\",T1.\"ItemCode\" as \"docItem\",T1.\"Dscription\" as \"docItemDes\",COALESCE(T1.\"ShipDate\", T0.\"DocDueDate\") AS \"docDelDate\",\n" +
                            "T1.\"LineTotal\" as \"docLineT\",T1.\"U_STXToolNum\" as \"ToolNum\",T1.\"U_STXPartNum\" as \"PartNum\",T1.\"U_STXPartName\" as \"PartName\",WO.\"DocEntry\" AS \"WO Entry\",\n" +
                            "WO.\"DocNum\" AS \"WONum\",TR.\"DocEntry\" AS \"Reception Entry\",TR.\"DocNum\" AS \"docToolNum\", T2.\"U_PartDsc\" as \"WorkDesc\"\n" +
                            "FROM ORDR T0\n" +
                            "INNER JOIN RDR1 T1 ON T0.\"DocEntry\" = T1.\"DocEntry\"\n" +
                            "INNER JOIN \"@STXQC19\" T2 ON T1.\"U_STXQC19ID\" = T2.\"DocEntry\"\n" +
                            "LEFT JOIN OWOR WO ON WO.\"U_STXSONum\" = T0.\"DocNum\" AND WO.\"U_STXSOLineNum\" = T1.\"LineNum\" AND WO.Status NOT IN ('C')\n" +
                            "LEFT JOIN OIGN TR ON TR.\"U_STXSONum\" = T0.\"DocNum\" AND TR.\"U_STXToolNum\" = T1.\"U_STXToolNum\"\n" +
                            "WHERE T0.\"DocNum\" = {0}";
            query = string.Format(query, openDocEntry);
            dataTable.ExecuteQuery(query);
        }


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

        internal static double VerifyCC2(string itemCC2)
        {
            double result = 0;

            string sSql = $"select \"PrcCode\" from OPRC where \"DimCode\" = 3 and \"PrcCode\" = '{itemCC2}'";
            Recordset rs = Utils.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset) as Recordset;
            rs.DoQuery(sSql);
            if (rs.RecordCount > 0)
            {
                result = 1;
            }
            else
            {
                result = 0;
            }

            return result;
        }

        internal static double VerifyCC1(string itemCC1)
        {
            double result = 0;

            string sSql = $"select \"PrcCode\" from OPRC where \"DimCode\" = 1 and \"PrcCode\" = '{itemCC1}'";
            Recordset rs = Utils.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset) as Recordset;
            rs.DoQuery(sSql);
            if (rs.RecordCount > 0)
            {
                result = 1;
            }
            else
            {
                result = 0;
            }

            return result;
        }

        internal static void CreateProductionOrder(string salesOrder, string lineNum)
        {

            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utils.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string queryHeader = "Select T0.\"DocNum\",T0.\"DocEntry\",T0.\"ObjType\",T1.\"LineNum\",T1.\"U_STXQC19ID\",T0.\"CardCode\",T0.\"CardName\",T0.\"LicTradNum\", T3.\"SlpName\",T5.\"U_Fin_Sta\",\n" +
                                "T0.\"U_STXBrand\",T0.\"U_STXOEMPgm\",T1.\"ItemCode\", T1.\"WhsCode\",T1.\"Quantity\",T1.\"OcrCode\",T1.\"OcrCode2\",T1.\"OcrCode3\",T1.\"OcrCode4\",T1.\"OcrCode5\"\n" +
                                "from ORDR T0\n" +
                                "inner join RDR1 T1 on T0.\"DocEntry\" = T1.\"DocEntry\"\n" +
                                "left join OSLP T3 on T0.\"SlpCode\" = T1.\"SlpCode\"\n" +
                                "left join OCRD T4 on T0.\"CardCode\" = T4.\"CardCode\"\n" +
                                "left join OACT T5 on T4.\"DebPayAcct\" = T5.\"AcctCode\"\n" +
                                "Where T0.\"DocNum\" = {0} and T1.\"LineNum\" = {1}";

            queryHeader = string.Format(queryHeader, salesOrder, lineNum);
            rs.DoQuery(queryHeader);

            while (!rs.EoF)
            {
                // Create a new Production Order object
                SAPbobsCOM.ProductionOrders prodOrder = (SAPbobsCOM.ProductionOrders)Utils.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders);

                // Set header properties
                prodOrder.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotSpecial;
                prodOrder.ProductionOrderStatus = BoProductionOrderStatusEnum.boposPlanned;
                if (rs.Fields.Item("U_Fin_Sta").Value.ToString() == "103")
                {
                    prodOrder.UserFields.Fields.Item("U_STXWOType").Value = "Intercompany Order";
                }
                else
                {
                    prodOrder.UserFields.Fields.Item("U_STXWOType").Value = "Normal Production";
                }
                
                prodOrder.ItemNo = rs.Fields.Item("ItemCode").Value.ToString();  // Assuming your query returns an ItemCode field
                prodOrder.PlannedQuantity = Convert.ToDouble(rs.Fields.Item("Quantity").Value);  // Assuming your query returns a Quantity field
                prodOrder.UserFields.Fields.Item("U_STXSONum").Value = Convert.ToInt32(rs.Fields.Item("DocNum").Value);
                prodOrder.UserFields.Fields.Item("U_STXSOLineNum").Value = Convert.ToInt32(rs.Fields.Item("LineNum").Value);
                prodOrder.UserFields.Fields.Item("U_STXQC19ID").Value = Convert.ToInt32(rs.Fields.Item("U_STXQC19ID").Value);
                prodOrder.Warehouse = rs.Fields.Item("WhsCode").Value.ToString();
                prodOrder.ProductionOrderOrigin = BoProductionOrderOriginEnum.bopooManual;
                prodOrder.DocumentReferences.ReferencedDocEntry = Convert.ToInt32(rs.Fields.Item("DocEntry").Value);
                prodOrder.DocumentReferences.ReferencedObjectType = (ReferencedObjectTypeEnum)Convert.ToInt32(rs.Fields.Item("ObjType").Value);

                prodOrder.CustomerCode = rs.Fields.Item("CardCode").Value.ToString();
                prodOrder.UserFields.Fields.Item("U_STXCustName").Value = rs.Fields.Item("CardName").Value;
                prodOrder.UserFields.Fields.Item("U_STXLicTradNum").Value = rs.Fields.Item("LicTradNum").Value.ToString();
                prodOrder.UserFields.Fields.Item("U_STXSalesEmployee").Value = rs.Fields.Item("SlpName").Value.ToString();
                prodOrder.UserFields.Fields.Item("U_STXWOBrand").Value = rs.Fields.Item("U_STXBrand").Value.ToString();
                prodOrder.UserFields.Fields.Item("U_STXOEMPgm").Value = rs.Fields.Item("U_STXOEMPgm").Value.ToString();
                prodOrder.DistributionRule = rs.Fields.Item("OcrCode").Value.ToString();
                prodOrder.DistributionRule2 = rs.Fields.Item("OcrCode2").Value.ToString();
                prodOrder.DistributionRule3 = rs.Fields.Item("OcrCode3").Value.ToString();
                prodOrder.DistributionRule4 = rs.Fields.Item("OcrCode4").Value.ToString();
                prodOrder.DistributionRule5 = rs.Fields.Item("OcrCode5").Value.ToString();

                

                SAPbobsCOM.Recordset lineRs = (SAPbobsCOM.Recordset)Utils.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string queryLines = "Select T2.\"VisOrder\",T2.\"U_Texture\",T2.\"U_resCode\",T2.\"U_opCode\",T2.\"U_opDesc\",T2.\"U_opDescL\",T1.\"WhsCode\",T2.\"U_Quantity\",T2.\"LineId\"\n" +
                                    ",T1.\"OcrCode\",T1.\"OcrCode2\",T1.\"OcrCode3\",T1.\"OcrCode4\",T1.\"OcrCode5\" from ORDR T0\n" +
                                    "inner join RDR1 T1 on T0.\"DocEntry\" = T1.\"DocEntry\"\n" +
                                    "inner join \"@STXQC19O\" T2 on T1.\"U_STXQC19ID\" = T2.\"DocEntry\"\n" +
                                    "Where T0.\"DocNum\" = {0} and T1.\"LineNum\" = {1}";

                queryLines = string.Format(queryLines, salesOrder, lineNum);
                lineRs.DoQuery(queryLines);
                while (!lineRs.EoF)
                {
                    // Add lines to the Production Order
                    //prodOrder.Lines.SetCurrentLine(prodOrder.Lines.Count);
                    prodOrder.Lines.ItemType = ProductionItemType.pit_Resource;
                    prodOrder.Lines.ItemNo = lineRs.Fields.Item("U_resCode").Value.ToString();
                    prodOrder.Lines.PlannedQuantity = Convert.ToDouble(lineRs.Fields.Item("U_Quantity").Value);
                    if (lineRs.Fields.Item("U_resCode").Value.ToString().StartsWith("SUB"))
                    {
                        prodOrder.Lines.Warehouse = "EXT";
                    }
                    else
                    {
                        prodOrder.Lines.Warehouse = lineRs.Fields.Item("WhsCode").Value.ToString();
                    }
                    
                    //prodOrder.Lines.UserFields.Fields.Item("U_Texture").Value = lineRs.Fields.Item("U_Texture").Value.ToString();

                    prodOrder.Lines.UserFields.Fields.Item("U_Texture").Value = lineRs.Fields.Item("U_Texture").Value.ToString();
                    prodOrder.Lines.UserFields.Fields.Item("U_STXOPCode").Value = lineRs.Fields.Item("U_opCode").Value.ToString();
                    prodOrder.Lines.UserFields.Fields.Item("U_STXOPDes").Value = lineRs.Fields.Item("U_opDesc").Value.ToString();
                    prodOrder.Lines.UserFields.Fields.Item("U_STXOPDesLocal").Value = lineRs.Fields.Item("U_opDescL").Value.ToString();
                    prodOrder.Lines.UserFields.Fields.Item("U_QCLineID").Value = lineRs.Fields.Item("LineId").Value.ToString();

                    prodOrder.Lines.DistributionRule = lineRs.Fields.Item("OcrCode").Value.ToString();
                    prodOrder.Lines.DistributionRule2 = lineRs.Fields.Item("OcrCode2").Value.ToString();
                    prodOrder.Lines.DistributionRule3 = lineRs.Fields.Item("OcrCode3").Value.ToString();
                    prodOrder.Lines.DistributionRule4 = lineRs.Fields.Item("OcrCode4").Value.ToString();
                    prodOrder.Lines.DistributionRule5 = lineRs.Fields.Item("OcrCode5").Value.ToString();

                    // Set other line-specific properties as required

                    prodOrder.Lines.Add();
                    lineRs.MoveNext();
                }

                // Add the Production Order
                int addResult = prodOrder.Add();
                if (addResult != 0)
                {
                    string errMsg = "";
                    int errCode = 0;
                    Utils.oCompany.GetLastError(out errCode, out errMsg);
                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox($"Error adding production order: {errCode} - {errMsg}");
                }
                else
                {
                    string newDocEntryStr = Utils.oCompany.GetNewObjectKey();
                    string newDocObjTypeStr = Utils.oCompany.GetNewObjectType();
                    int newDocEntry = int.Parse(newDocEntryStr);
                    int newObjType = int.Parse(newDocObjTypeStr);
                    int baseEntry = Convert.ToInt32(rs.Fields.Item("DocEntry").Value);
                    UpdateSalesOrderReference(baseEntry, newDocEntry, newObjType);
                }
                

                rs.MoveNext();
            }
        }

        private static bool UpdateSalesOrderReference(int docEntry, int referencedDocEntry, int referencedObjectType)
        {
            SAPbobsCOM.Documents salesOrder = (SAPbobsCOM.Documents)Utils.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);

            if (salesOrder.GetByKey(docEntry)) // Assuming docEntry is the DocEntry of the Sales Order you wish to modify
            {

                // Add new reference if not found
                salesOrder.DocumentReferences.Add();
                salesOrder.DocumentReferences.ReferencedDocEntry = referencedDocEntry;
                salesOrder.DocumentReferences.ReferencedObjectType = (SAPbobsCOM.ReferencedObjectTypeEnum)referencedObjectType;

                // Commit the changes
                int result = salesOrder.Update();
                if (result == 0)
                    return true;
                else
                {
                    int errCode;
                    string errMsg;
                    Utils.oCompany.GetLastError(out errCode, out errMsg);
                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox($"Error updating sales order: {errCode} - {errMsg}");
                    return false;
                }
            }
            else
            {
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox($"Sales Order with DocEntry {docEntry} not found.");
                return false;
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

        internal static string GetOperation(SAPbouiCOM.DataTable operations, IForm uIAPIRawForm, Matrix mOperations, string CalcFactor, string concatenatedTextureCodes, string tclassFactor, string OpQuantityExpression, string SptCode, bool DefBOM,string QtyFactorExpression)
        {
            string query = "";
            if (((SAPbouiCOM.CheckBox)uIAPIRawForm.Items.Item("DefBOM").Specific).Checked == true)
            {

                query = "WITH ResourceCosts AS(SELECT \"ResCode\",\"ResName\",\"UnitOfMsr\",(COALESCE(\"StdCost1\",0)+COALESCE(\"StdCost2\",0)+COALESCE(\"StdCost3\",0)+COALESCE(\"StdCost4\",0)+COALESCE(\"StdCost5\",0)+COALESCE(\"StdCost6\",0)+COALESCE(\"StdCost7\",0)+COALESCE(\"StdCost8\",0)+COALESCE(\"StdCost9\",0) +COALESCE(\"StdCost10\",0)) AS \"ResCost\" FROM ORSC),\n" +

                "BaseData AS(\n" +
                "select CASE WHEN T3.\"U_PlanType\" = 'I' THEN 1 WHEN T3.\"U_PlanType\" NOT IN ('I', 'F') THEN 2 WHEN T3.\"U_PlanType\" = 'F' THEN 3 END as \"Order\",T1.\"VisOrder\" as \"U_groupOrder\", \n" +
                "T1.\"VisOrder\" as \"U_operationOrder\" , T3.\"U_PlanType\", '' as \"Texture\",\n" +
                "T1.\"Code\" as \"U_operationResource\", T1.\"U_STXOPCode\" as \"U_operationCode\", T3.\"U_STXOPDes\", T3.\"U_STXOPDesLocal\",1 as \"CalcFactor\",\n" +
                "T1.\"U_STXQtyBy\",T0.\"PlAvgSize\",T1.\"Quantity\" as \"Quantity\",1 as \"TClassFactor\",1 as \"QtyFactor\"\n" +
                "from OITT T0\n" +
                "inner join ITT1 T1 on T0.\"Code\" = T1.\"Father\"\n" +
                "left join \"@STXOPERATIONS\" T3 on T1.\"U_STXOPCode\"= T3.\"Code\"\n" +
                "where T0.\"Code\" = '{2}'),\n" +

                "CalculatedQuantities AS (SELECT *,CASE WHEN \"U_STXQtyBy\" = 'A' THEN \"QtyFactor\" * \"Quantity\" * (\"CalcFactor\" / \"PlAvgSize\") * \"TClassFactor\" * \"TClassFactor\" ELSE \"QtyFactor\" * \"Quantity\" * \"TClassFactor\" END AS \"CalculatedQty\" FROM BaseData),\n" +
                "MaxQuantities AS(SELECT \"Order\", \"U_operationOrder\", \"U_PlanType\", \"U_operationResource\", \"U_operationCode\",MAX(\"CalculatedQty\") AS \"MaxCalculatedQty\" FROM CalculatedQuantities Where U_PlanType != 'N' GROUP BY \"Order\", U_operationOrder, U_PlanType, U_operationResource, U_operationCode),\n" +
                "FilteredOperations as (SELECT T1.*FROM MaxQuantities T0 JOIN CalculatedQuantities T1 ON T0.\"Order\" = T1.\"Order\" AND T0.\"U_operationOrder\" = T1.\"U_operationOrder\" AND T0.\"U_PlanType\" = T1.\"U_PlanType\" AND\n" +
                "T0.\"U_operationResource\" = T1.\"U_operationResource\" AND T0.\"U_operationCode\" = T1.\"U_operationCode\" AND T0.\"MaxCalculatedQty\" = T1.\"CalculatedQty\")\n" +

                "select ROW_NUMBER() OVER(ORDER BY X0.\"Order\",X0.\"Texture\",X0.\"U_groupOrder\",X0.\"U_operationOrder\") AS \"VisOrder\",X0.\"Texture\" as \"U_Texture\",X0.\"U_operationResource\" as \"U_resCode\",X0.\"ResName\" as \"U_resName\",X0.\"U_operationCode\" as \"U_opCode\",\n" +
                "X0.\"U_STXOPDes\" as \"U_opDesc\",X0.\"U_STXOPDesLocal\" as \"U_opDescL\",CONVERT(nvarchar,cast(Round(X0.\"QTY\",{5}) AS DECIMAL(18, {5}))) as \"U_sugQty\",\n" +
                "CONVERT(nvarchar, cast(Round(X0.\"QTY\",{5}) AS DECIMAL(18, {5}))) as \"U_Quantity\",X0.\"UnitOfMsr\" as \"U_UOM\",CONVERT(nvarchar,cast(Round((X0.\"ResCost\"),{6}) AS DECIMAL(18, {6}))) as \"U_Price\",\n" +
                "CONVERT(nvarchar, cast(Round(X0.\"QTY\" * X0.\"ResCost\",{6}) AS DECIMAL(18, {6}))) as \"U_LineTot\",\n" +
                "Case when coalesce(isnull(X0.\"ResName\",''),'') = '' then '{8}' when coalesce(isnull(X0.\"U_STXOPDes\",''),'') = '' then '{9}' when(X0.\"U_STXQtyBy\" = 'A' and X0.\"CalcFactor\" = 0) then '{10}' when X0.\"ResCost\" = 0 then '{11}' end as \"U_ErrMsg\",\n" +
                "Case when X0.\"U_PlanType\" = 'I' then 0 when  X0.\"U_PlanType\" = 'F' then 99 else DENSE_RANK() OVER (order by X0.\"Texture\")-1 end as \"U_seq\" from (\n" +
                "Select R0.\"Order\",CASE WHEN R0.\"U_PlanType\" = 'N' then R0.\"U_groupOrder\" else NULL END as \"U_groupOrder\",R0.\"U_operationOrder\", R0.\"U_PlanType\",R0.\"Texture\",R0.\"U_operationResource\",R1.\"ResName\",R0.\"U_operationCode\",\n" +
                "R0.\"U_STXOPDes\",R0.\"U_STXOPDesLocal\",R0.\"PlAvgSize\",R0.\"Quantity\" as \"Quantity\",R0.\"U_STXQtyBy\",R0.QtyFactor,R0.\"CalcFactor\",R0.TClassFactor,R1.\"ResCost\",R1.\"UnitOfMsr\"\n" +
                ",Case When U_STXQtyBy = 'A' then R0.QtyFactor* R0.Quantity * (R0.CalcFactor / R0.PlAvgSize) * R0.TClassFactor else R0.QtyFactor* R0.Quantity* R0.TClassFactor end as \"Qty\"\n" +
                "from BaseData R0\n" +
                "left join ResourceCosts R1 on R0.\"U_operationResource\" = R1.\"ResCode\"\n" +
                "where R0.\"U_PlanType\" = 'N'\n" +
                "union all\n" +
                "SELECT distinct R0.\"Order\",CASE WHEN R0.\"U_PlanType\" = 'N' then R0.\"U_groupOrder\" else NULL END as \"U_groupOrder\",R0.\"U_operationOrder\", R0.\"U_PlanType\",R0.\"Texture\",R0.\"U_operationResource\",R1.\"ResName\",R0.\"U_operationCode\",\n" +
                "R0.\"U_STXOPDes\",R0.\"U_STXOPDesLocal\",R0.\"PlAvgSize\",R0.\"Quantity\" as \"Quantity\",R0.\"U_STXQtyBy\",R0.QtyFactor,R0.\"CalcFactor\",R0.TClassFactor,R1.\"ResCost\",R1.\"UnitOfMsr\"\n" +
                ",R0.\"CalculatedQty\" as \"Qty\"\n" +
                "FROM FilteredOperations R0\n" +
                "LEFT JOIN ResourceCosts R1 ON R0.\"U_operationResource\" = R1.\"ResCode\") X0, OADM X1 \n" +
                "order by X0.\"Order\",X0.\"Texture\",X0.\"U_groupOrder\",X0.\"U_operationOrder\"";

                query = string.Format(query, CalcFactor, concatenatedTextureCodes, SptCode, tclassFactor, OpQuantityExpression, Utils.QtyDec, Utils.PriceDec, Utils.SumDec, Resources.mOperErr1, Resources.mOperErr2, Resources.mOperErr3, Resources.mOperErr4, DefBOM, QtyFactorExpression);

            }
            else
            {
                query = "WITH TextureTasks AS (SELECT * FROM \"@STXSETPTEXTURETASKS\" WHERE \"U_standexReference\" IN({1})),\n" +

                "ResourceCosts AS(SELECT \"ResCode\",\"ResName\",\"UnitOfMsr\",(COALESCE(\"StdCost1\",0)+COALESCE(\"StdCost2\",0)+COALESCE(\"StdCost3\",0)+COALESCE(\"StdCost4\",0)+COALESCE(\"StdCost5\",0)+COALESCE(\"StdCost6\",0)+COALESCE(\"StdCost7\",0)+COALESCE(\"StdCost8\",0)+COALESCE(\"StdCost9\",0) +COALESCE(\"StdCost10\",0)) AS \"ResCost\" FROM ORSC),\n" +

                "BaseData AS(\n" +
                "select CASE WHEN T3.\"U_PlanType\" = 'I' THEN 1 WHEN T3.\"U_PlanType\" NOT IN ('I', 'F') THEN 2 WHEN T3.\"U_PlanType\" = 'F' THEN 3 END as \"Order\",T2.\"U_groupOrder\", \n" +
                "Case When '{12}' = 'True' then T1.\"VisOrder\" else T2.\"U_operationOrder\" end as \"U_operationOrder\" , T3.\"U_PlanType\", Case When T3.\"U_PlanType\" IN ('I', 'F') THEN NULL else T2.\"U_standexReference\" end as \"Texture\",\n" +
                "T2.\"U_operationResource\", T2.\"U_operationCode\", T3.\"U_STXOPDes\", T3.\"U_STXOPDesLocal\",{0},\n" +
                "T1.\"U_STXQtyBy\",T0.\"PlAvgSize\",{4},{3},{13},T2.\"U_standexReference\"\n" +
                "from OITT T0\n" +
                "inner join ITT1 T1 on T0.\"Code\" = T1.\"Father\"\n" +
                "inner join TextureTasks T2 ON T1.\"U_STXOPCode\" = T2.\"U_operationCode\"\n" +
                "left join \"@STXOPERATIONS\" T3 on T2.\"U_operationCode\"= T3.\"Code\"\n" +
                "where T0.\"Code\" = '{2}'),\n" +

                "CalculatedQuantities AS (SELECT *,CASE WHEN \"U_STXQtyBy\" = 'A' THEN \"QtyFactor\" * \"Quantity\" * (\"CalcFactor\" / \"PlAvgSize\") * \"TClassFactor\" * \"TClassFactor\" ELSE \"QtyFactor\" * \"Quantity\" * \"TClassFactor\" END AS \"CalculatedQty\" FROM BaseData),\n" +
                "MaxQuantities AS(SELECT \"Order\", \"U_operationOrder\", \"U_PlanType\", \"U_operationResource\", \"U_operationCode\",MAX(\"CalculatedQty\") AS \"MaxCalculatedQty\" FROM CalculatedQuantities Where U_PlanType != 'N' GROUP BY \"Order\", U_operationOrder, U_PlanType, U_operationResource, U_operationCode),\n" +
                "FilteredOperations as (SELECT T1.*FROM MaxQuantities T0 JOIN CalculatedQuantities T1 ON T0.\"Order\" = T1.\"Order\" AND T0.\"U_operationOrder\" = T1.\"U_operationOrder\" AND T0.\"U_PlanType\" = T1.\"U_PlanType\" AND\n" +
                "T0.\"U_operationResource\" = T1.\"U_operationResource\" AND T0.\"U_operationCode\" = T1.\"U_operationCode\" AND T0.\"MaxCalculatedQty\" = T1.\"CalculatedQty\")\n" +

                "select ROW_NUMBER() OVER(ORDER BY X0.\"Order\",X0.\"Texture\",X0.\"U_groupOrder\",X0.\"U_operationOrder\") AS \"VisOrder\",X0.\"Texture\" as \"U_Texture\",X0.\"U_operationResource\" as \"U_resCode\",X0.\"ResName\" as \"U_resName\",X0.\"U_operationCode\" as \"U_opCode\",\n" +
                "X0.\"U_STXOPDes\" as \"U_opDesc\",X0.\"U_STXOPDesLocal\" as \"U_opDescL\",CONVERT(nvarchar,cast(Round(X0.\"QTY\",{5}) AS DECIMAL(18, {5}))) as \"U_sugQty\",\n" +
                "CONVERT(nvarchar, cast(Round(X0.\"QTY\",{5}) AS DECIMAL(18, {5}))) as \"U_Quantity\",X0.\"UnitOfMsr\" as \"U_UOM\",CONVERT(nvarchar,cast(Round((X0.\"ResCost\"),{6}) AS DECIMAL(18, {6}))) as \"U_Price\",\n" +
                "CONVERT(nvarchar, cast(Round(X0.\"QTY\" * X0.\"ResCost\",{6}) AS DECIMAL(18, {6}))) as \"U_LineTot\",\n" +
                "Case when coalesce(isnull(X0.\"ResName\",''),'') = '' then '{8}' when coalesce(isnull(X0.\"U_STXOPDes\",''),'') = '' then '{9}' when(X0.\"U_STXQtyBy\" = 'A' and X0.\"CalcFactor\" = 0) then '{10}' when X0.\"ResCost\" = 0 then '{11}' end as \"U_ErrMsg\",\n" +
                "Case when X0.\"U_PlanType\" = 'I' then 0 when  X0.\"U_PlanType\" = 'F' then 99 else DENSE_RANK() OVER (order by X0.\"Texture\")-1 end as \"U_seq\" from (\n" +
                "Select R0.\"Order\",CASE WHEN R0.\"U_PlanType\" = 'N' then R0.\"U_groupOrder\" else NULL END as \"U_groupOrder\",R0.\"U_operationOrder\", R0.\"U_PlanType\",R0.\"Texture\",R0.\"U_operationResource\",R1.\"ResName\",R0.\"U_operationCode\",\n" +
                "R0.\"U_STXOPDes\",R0.\"U_STXOPDesLocal\",R0.\"PlAvgSize\",R0.\"Quantity\" as \"Quantity\",R0.\"U_STXQtyBy\",R0.QtyFactor,R0.\"CalcFactor\",R0.TClassFactor,R1.\"ResCost\",R1.\"UnitOfMsr\"\n" +
                ",Case When U_STXQtyBy = 'A' then R0.QtyFactor* R0.Quantity * (R0.CalcFactor / R0.PlAvgSize) * R0.TClassFactor else R0.QtyFactor* R0.Quantity* R0.TClassFactor end as \"Qty\"\n" +
                "from BaseData R0\n" +
                "left join ResourceCosts R1 on R0.\"U_operationResource\" = R1.\"ResCode\"\n" +
                "where R0.\"U_PlanType\" = 'N'\n" +
                "union all\n" +
                "SELECT distinct R0.\"Order\",CASE WHEN R0.\"U_PlanType\" = 'N' then R0.\"U_groupOrder\" else NULL END as \"U_groupOrder\",R0.\"U_operationOrder\", R0.\"U_PlanType\",R0.\"Texture\",R0.\"U_operationResource\",R1.\"ResName\",R0.\"U_operationCode\",\n" +
                "R0.\"U_STXOPDes\",R0.\"U_STXOPDesLocal\",R0.\"PlAvgSize\",R0.\"Quantity\" as \"Quantity\",R0.\"U_STXQtyBy\",R0.QtyFactor,R0.\"CalcFactor\",R0.TClassFactor,R1.\"ResCost\",R1.\"UnitOfMsr\"\n" +
                ",R0.\"CalculatedQty\" as \"Qty\"\n" +
                "FROM FilteredOperations R0\n" +
                "LEFT JOIN ResourceCosts R1 ON R0.\"U_operationResource\" = R1.\"ResCode\") X0, OADM X1 \n" +
                "order by X0.\"Order\",X0.\"Texture\",X0.\"U_groupOrder\",X0.\"U_operationOrder\"";

                query = string.Format(query, CalcFactor, concatenatedTextureCodes, SptCode, tclassFactor, OpQuantityExpression, Utils.QtyDec, Utils.PriceDec, Utils.SumDec, Resources.mOperErr1, Resources.mOperErr2, Resources.mOperErr3, Resources.mOperErr4, DefBOM,QtyFactorExpression);

            }

            try
            {
                operations.ExecuteQuery(query);

                string xml = operations.SerializeAsXML(BoDataTableXmlSelect.dxs_All);
                TempDataTable mdt = xml.XmlDeserializeFromString<TempDataTable>();

                string xmlOperations = (XMLDatasource.GenerateXml(mdt)).ToString();
                return xmlOperations;

            }
            catch (Exception ex)
            {

                Program.SBO_Application.SetStatusBarMessage(ex.Message);
                return null;
            }
        }

        internal static void UpdateQCIDBaseDoc(string qcidValue, string sapdocEntry, string value)
        {
            SAPbobsCOM.CompanyService oCompanyService = Utils.oCompany.GetCompanyService();
            SAPbobsCOM.GeneralService oGeneralService = oCompanyService.GetGeneralService("STXQC19");


            SAPbobsCOM.GeneralDataParams oParameters = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
            oParameters.SetProperty("DocEntry", qcidValue);
            SAPbobsCOM.GeneralData oldEntry = oGeneralService.GetByParams(oParameters);
            oldEntry.SetProperty("U_bsDocEntry", sapdocEntry);

            oGeneralService.Update(oldEntry);
        }

        internal static int getDocLineofQCID(string qcidValue, string sapdocEntry, string sapObjType)
        {
            int docLine = -1;
            string sSql = $"select \"U_bsDocEntry\",\"U_bsLineNum\",\"U_bsObjType\" from \"@STXQC19\" where \"DocEntry\" = '{qcidValue}' and \"U_bsDocEntry\" = '{sapdocEntry}' and \"U_bsObjType\" = '{sapObjType}'";
            Recordset rs = Utils.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset) as Recordset;
            rs.DoQuery(sSql);

            if (!rs.EoF)
            {
                docLine = (int)rs.Fields.Item("U_bsLineNum").Value;
            }

            return docLine;
        }

        internal static string duplicateQCID(string qcidValue, string sapdocEntry, string sapObjType, string intLineNo, bool itmChange)
        {
            string sapDocE = "";
            string sapObjTyp = "";
            string sapLineNo = "";

            string sSql = $"select \"DocEntry\",\"U_bsDocEntry\",\"U_bsObjType\",\"U_bsLineNum\" from \"@STXQC19\" where \"DocEntry\" = '{qcidValue}'";
            Recordset rs = Utils.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset) as Recordset;
            rs.DoQuery(sSql);

            if (!rs.EoF)
            {
                sapDocE = rs.Fields.Item("U_bsDocEntry").Value.ToString();
                sapObjTyp = rs.Fields.Item("U_bsObjType").Value.ToString() ;
                sapLineNo = rs.Fields.Item("U_bsLineNum").Value.ToString();

            }

            if (rs.RecordCount > 0 && (sapdocEntry != sapDocE || sapObjTyp != sapObjType))
            {
                    SAPbobsCOM.CompanyService oCompanyService = Utils.oCompany.GetCompanyService();
                    SAPbobsCOM.GeneralService oGeneralService = oCompanyService.GetGeneralService("STXQC19");


                    SAPbobsCOM.GeneralDataParams oParameters = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                    oParameters.SetProperty("DocEntry", qcidValue);

                    // Get the UDO entry you wish to duplicate
                    SAPbobsCOM.GeneralData oldEntry = oGeneralService.GetByParams(oParameters);
                    oldEntry.SetProperty("U_bsLineNum", intLineNo);
                    oldEntry.SetProperty("U_bsDocEntry", sapdocEntry);
                    oldEntry.SetProperty("U_bsObjType", sapObjType);

                    SAPbobsCOM.GeneralDataParams newEntryParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.Add(oldEntry);

                    // Get the DocEntry of the newly added record
                    string newEntryNumber = newEntryParams.GetProperty("DocEntry").ToString();

                    return newEntryNumber;
            }

            else if (rs.RecordCount > 0 && (sapdocEntry == sapDocE && sapObjTyp == sapObjType) && intLineNo != sapLineNo)
            {
                if (itmChange == true)
                {
                    SAPbobsCOM.CompanyService oCompanyService = Utils.oCompany.GetCompanyService();
                    SAPbobsCOM.GeneralService oGeneralService = oCompanyService.GetGeneralService("STXQC19");


                    SAPbobsCOM.GeneralDataParams oParameters = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                    oParameters.SetProperty("DocEntry", qcidValue);

                    // Get the UDO entry you wish to duplicate
                    SAPbobsCOM.GeneralData oldEntry = oGeneralService.GetByParams(oParameters);


                    // Remove or modify the specific child table data
                    SAPbobsCOM.GeneralDataCollection childTable = oldEntry.Child("STXQC19O"); 
                                                                                                               
                    while (childTable.Count > 0)
                    {
                        childTable.Remove(0); 
                    }

                    oldEntry.SetProperty("U_bsLineNum", intLineNo);
                    oldEntry.SetProperty("U_bsDocEntry", sapdocEntry);
                    oldEntry.SetProperty("U_bsObjType", sapObjType);

                    SAPbobsCOM.GeneralDataParams newEntryParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.Add(oldEntry);

                    // Get the DocEntry of the newly added record
                    string newEntryNumber = newEntryParams.GetProperty("DocEntry").ToString();

                    return newEntryNumber;
                }
                else
                {
                    SAPbobsCOM.CompanyService oCompanyService = Utils.oCompany.GetCompanyService();
                    SAPbobsCOM.GeneralService oGeneralService = oCompanyService.GetGeneralService("STXQC19");


                    SAPbobsCOM.GeneralDataParams oParameters = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                    oParameters.SetProperty("DocEntry", qcidValue);

                    // Get the UDO entry you wish to duplicate
                    SAPbobsCOM.GeneralData oldEntry = oGeneralService.GetByParams(oParameters);
                    oldEntry.SetProperty("U_bsLineNum", intLineNo);
                    oldEntry.SetProperty("U_bsDocEntry", sapdocEntry);
                    oldEntry.SetProperty("U_bsObjType", sapObjType);

                    SAPbobsCOM.GeneralDataParams newEntryParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.Add(oldEntry);


                    // Get the DocEntry of the newly added record
                    string newEntryNumber = newEntryParams.GetProperty("DocEntry").ToString();

                    return newEntryNumber;
                }
            }
            else
            {
                if (string.IsNullOrEmpty(qcidValue))
                {
                        return null;
                }
                else
                {
                    return qcidValue;
                }
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

        //internal static void UpdateOperationsDB(System.Data.DataTable mOperations, string qCDocEntry)
        //{
        //    QCEvents.operationsUpdate = false;

        //    SAPbobsCOM.GeneralData oChild = null;
        //    SAPbobsCOM.GeneralDataCollection oChildren = null;

        //    //SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)uIAPIRawForm.Items.Item("mOper").Specific;

        //    SAPbobsCOM.CompanyService oCompanyService = Utils.oCompany.GetCompanyService();
        //    SAPbobsCOM.GeneralService oGeneralService = oCompanyService.GetGeneralService("STXQC19");
        //    SAPbobsCOM.GeneralData oGeneralData = (SAPbobsCOM.GeneralData)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
        //    SAPbobsCOM.GeneralDataParams oGeneralParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);


        //    for (int i = 0; i < mOperations.Rows.Count; i++)
        //    {
        //        try
        //        {
        //            oGeneralParams.SetProperty("DocEntry", qCDocEntry);      //Primary Key
        //            oGeneralData = oGeneralService.GetByParams(oGeneralParams);

        //            oChildren = oGeneralData.Child("STXQC19O");

        //            // Check if the child at index i exists
        //            if (i < oChildren.Count)
        //            {
        //                // If it exists, retrieve it
        //                oChild = oChildren.Item(i);
        //            }
        //            else
        //            {
        //                // If it doesn't exist, add a new child and then retrieve it
        //                oChildren.Add();
        //                oChild = oChildren.Item(oChildren.Count - 1);
        //            }

        //            oChild.SetProperty("U_Texture", mOperations.Rows[i]["OPTexture"]);
        //            oChild.SetProperty("U_resCode", mOperations.Rows[i]["OPResc"]);
        //            oChild.SetProperty("U_resName", mOperations.Rows[i]["OPResN"]);
        //            oChild.SetProperty("U_opCode", mOperations.Rows[i]["OPcode"]);
        //            oChild.SetProperty("U_opDesc", mOperations.Rows[i]["OPName"]);
        //            oChild.SetProperty("U_opDescL", mOperations.Rows[i]["OPNameL"]);
        //            oChild.SetProperty("U_sugQty", mOperations.Rows[i]["OPStdT"]);
        //            oChild.SetProperty("U_Quantity", mOperations.Rows[i]["OPQtdT"]);
        //            oChild.SetProperty("U_UOM", mOperations.Rows[i]["OPUom"]);
        //            oChild.SetProperty("U_Price", mOperations.Rows[i]["OPCost"]);
        //            oChild.SetProperty("U_LineTot", mOperations.Rows[i]["OPTotal"]);
        //            oChild.SetProperty("U_ErrMsg", mOperations.Rows[i]["OPErrMsg"]);
        //            oChild.SetProperty("U_seq", mOperations.Rows[i]["OPSeq"]);

        //            //Update the UDO Record                
        //            oGeneralService.Update(oGeneralData);   // If Child Table does not have any record, it will create; else, update the existing one

        //        }
        //        catch (Exception ex)
        //        {
        //            Program.SBO_Application.SetStatusBarMessage(ex.Message, BoMessageTime.bmt_Medium, true);
        //        }
        //    }
        //    for (int j = oChildren.Count - 1; j >= mOperations.Rows.Count; j--)
        //    {
        //        oChildren.Remove(j);
        //        oGeneralService.Update(oGeneralData);
        //    }
        //    //Program.SBO_Application.SetStatusBarMessage("Operations imported sucessfully.", BoMessageTime.bmt_Medium, false);
        //}

        internal static (string, string) GetSPT(SAPbouiCOM.EditText qCSubPart)
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

            string newQCSubPart = qCSubPart.String.Length > 2 ? qCSubPart.String.Substring(0, qCSubPart.String.Length - 2) + "00" : qCSubPart.String + "00";
            string sSql2 = $"SELECT T0.\"ItemCode\", T0.\"ItemName\" as \"Part Name\" FROM OITM T0 WHERE T0.\"ItemCode\" like '{newQCSubPart}'";

            Recordset rs2 = Utils.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset) as Recordset;
            rs2.DoQuery(sSql2);

            if (!rs2.EoF)
            {
                spt = (string)rs2.Fields.Item("ItemCode").Value;
                descr = descr + (string)rs2.Fields.Item("Part Name").Value;
            }
            return (spt, descr);
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

        internal static BoObjectTypes GetSAPObjectType(string sapObj)
        {
            // Create a dictionary to map numeric values to BoObjectTypes
            Dictionary<string, BoObjectTypes> objectMapping = new Dictionary<string, BoObjectTypes>
            {
                { "23", BoObjectTypes.oQuotations },
                { "17", BoObjectTypes.oOrders },
                { "202", BoObjectTypes.oWorkOrders }
            };

            // Check if the sapObj value is in the dictionary
            if (objectMapping.TryGetValue(sapObj, out BoObjectTypes objectType))
            {
                return objectType;
            }
            else
            {
                return BoObjectTypes.oQuotations; // Return a default value when no match is found
            }
        }

        internal static string GetObjectTypeCodeByFormType(string formType)
        {
            switch (formType)
            {
                case "149":
                    return "OQUT";  // Quotations
                case "139":
                    return "ORDR";  // Sales Orders
        
                default:
                    return string.Empty;
            }
        }

        internal static void UpdateSAPDocument(QuoteCalculator.QCResults unloadResults,string sapDocEntry,string sapObjType,string sapDocLineNum)
        {
            System.Globalization.NumberFormatInfo sapNumberFormat = Utils.GetSAPNumberFormatInfo();

            int docLine = -1;
            int docEntry = -1;
            int visOrder = -1;
            int? baseEntry = null;
            string SapObj = "-1";
            string tooln = "";string pName = "";string pNum = "";
            string doctooln = "";string docpName = "";string docpNum = "";

            string sSql = $"select \"U_bsDocEntry\",\"U_bsLineNum\",\"U_bsObjType\",\"U_ToolNum\",\"U_PartNum\",\"U_PartName\" from \"@STXQC19\" where \"DocEntry\" = '{unloadResults.QCID}'";
            Recordset rs = Utils.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset) as Recordset;
            rs.DoQuery(sSql);

            if (!rs.EoF)
            {
                docEntry = (int)rs.Fields.Item("U_bsDocEntry").Value;
                docLine = (int)rs.Fields.Item("U_bsLineNum").Value;
                SapObj = (string)rs.Fields.Item("U_bsObjType").Value;

                tooln = (string)rs.Fields.Item("U_ToolNum").Value;
                pName = (string)rs.Fields.Item("U_PartName").Value;
                pNum = (string)rs.Fields.Item("U_PartNum").Value;
            }

            if (SapObj == sapObjType && docEntry == int.Parse(sapDocEntry))
            {
                BoObjectTypes objectType = GetSAPObjectType(SapObj);
                string strObjType = GetSAPObjectLineStr(objectType);

                string sSql2 = $"select \"VisOrder\",\"BaseEntry\",\"U_STXToolNum\",\"U_STXPartNum\",\"U_STXPartName\",\"U_STXLeadTime\" from {strObjType} where DocEntry = '{docEntry}' and LineNum = '{docLine}'  and U_STXQC19ID = '{unloadResults.QCID}'";
                Recordset rs2 = Utils.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset) as Recordset;
                rs2.DoQuery(sSql2);

                if (!rs2.EoF)
                {
                    visOrder = (int)rs2.Fields.Item("VisOrder").Value;
                    baseEntry = (int)rs2.Fields.Item("BaseEntry").Value;

                    doctooln = (string)rs2.Fields.Item("U_STXToolNum").Value;
                    docpName = (string)rs2.Fields.Item("U_STXPartName").Value;
                    docpNum = (string)rs2.Fields.Item("U_STXPartNum").Value;
                }

                if (pName == unloadResults.QCprtName && pNum == unloadResults.QCptNum && tooln == unloadResults.QCtNum && baseEntry == 0)
                {
                    Documents sapDoc = (Documents)Utils.oCompany.GetBusinessObject(objectType);
                    if (sapDoc.GetByKey(docEntry))
                    {
                        sapDoc.Lines.SetCurrentLine(visOrder); // SetCurrentLine is 0-based
                        sapDoc.Lines.UnitPrice = HelperMethods.ParseDoubleWCur(unloadResults.QCuPrice, sapNumberFormat);
                        sapDoc.Lines.GrossBuyPrice = HelperMethods.ParseDoubleWCur(unloadResults.QCcPrice, sapNumberFormat);
                        if (doctooln != unloadResults.QCtNum)
                        {
                            sapDoc.Lines.UserFields.Fields.Item("U_STXToolNum").Value = unloadResults.QCtNum;
                        }
                        if (docpNum != unloadResults.QCptNum)
                        {
                            sapDoc.Lines.UserFields.Fields.Item("U_STXPartNum").Value = unloadResults.QCptNum;
                        }
                        if (docpName != unloadResults.QCprtName)
                        {
                            sapDoc.Lines.UserFields.Fields.Item("U_STXPartName").Value = unloadResults.QCprtName;
                        }

                        sapDoc.Lines.UserFields.Fields.Item("U_STXLeadTime").Value = unloadResults.QClTime;

                        if (sapDoc.Update() == 0)
                        {
                            // Update successful
                        }
                        else
                        {
                            string errorMessage = "";
                            Program.SBO_Application.MessageBox(errorMessage);
                        }
                    }
                }
                else
                {
                    // add new qcid entry and update the document
                }
            }
            else
            {
                if (pName == unloadResults.QCprtName && pNum == unloadResults.QCptNum && tooln == unloadResults.QCtNum)
                {
                    BoObjectTypes objectType = GetSAPObjectType(sapObjType);
                    string strObjType = GetSAPObjectLineStr(objectType);

                    string sSql2 = $"select \"VisOrder\",\"U_STXToolNum\",\"U_STXPartNum\",\"U_STXPartName\",\"U_STXLeadTime\" from {strObjType} where DocEntry = '{sapDocEntry}' and LineNum = '{sapDocLineNum}'  and U_STXQC19ID = '{unloadResults.QCID}'";
                    Recordset rs2 = Utils.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset) as Recordset;
                    rs2.DoQuery(sSql2);

                    if (!rs2.EoF)
                    {
                        visOrder = (int)rs2.Fields.Item("VisOrder").Value;
                        doctooln = (string)rs2.Fields.Item("U_STXToolNum").Value;
                        docpName = (string)rs2.Fields.Item("U_STXPartName").Value;
                        docpNum = (string)rs2.Fields.Item("U_STXPartNum").Value;
                    }

                    Documents sapDoc = (Documents)Utils.oCompany.GetBusinessObject(objectType);
                    if (sapDoc.GetByKey(int.Parse(sapDocEntry)))
                    {
                        sapDoc.Lines.SetCurrentLine(visOrder); // SetCurrentLine is 0-based
                        sapDoc.Lines.UnitPrice = HelperMethods.ParseDoubleWCur(unloadResults.QCuPrice, sapNumberFormat);
                        sapDoc.Lines.GrossBuyPrice = HelperMethods.ParseDoubleWCur(unloadResults.QCcPrice, sapNumberFormat);
                        if (doctooln != unloadResults.QCtNum)
                        {
                            sapDoc.Lines.UserFields.Fields.Item("U_STXToolNum").Value = unloadResults.QCtNum;
                        }
                        if (docpNum != unloadResults.QCptNum)
                        {
                            sapDoc.Lines.UserFields.Fields.Item("U_STXPartNum").Value = unloadResults.QCptNum;
                        }
                        if (docpName != unloadResults.QCprtName)
                        {
                            sapDoc.Lines.UserFields.Fields.Item("U_STXPartName").Value = unloadResults.QCprtName;
                        }
                        sapDoc.Lines.UserFields.Fields.Item("U_STXLeadTime").Value = unloadResults.QClTime;

                        if (sapDoc.Update() == 0)
                        {
                            // Update successful
                        }
                        else
                        {
                            string errorMessage = "";
                            Program.SBO_Application.MessageBox(errorMessage);
                        }
                    }
                }
                else
                {
                    // add new qcid entry and update the document
                }
            }
        }

        private static string GetSAPObjectLineStr(BoObjectTypes objectType)
        {
            // Create a dictionary to map BoObjectTypes to their string representations
            Dictionary<BoObjectTypes, string> objectMapping = new Dictionary<BoObjectTypes, string>
            {
                { BoObjectTypes.oQuotations, "QUT1" },
                { BoObjectTypes.oOrders, "RDR1" },
                { BoObjectTypes.oWorkOrders, "WOR1" }
            };

            // Use LINQ to find the key based on its value
            var matchedType = objectMapping.FirstOrDefault(x => x.Key == objectType);

            // If found, return the string value, else return default
            return matchedType.Equals(default(KeyValuePair<BoObjectTypes, string>)) ? "QUT1" : matchedType.Value;  // Assuming "23" (oQuotations) as default
        }


        internal static (string sMkSeg1Name,string sMkseg1ID, string sBrandName,string sBrandID, string sOEM, string sOEMProgram, string sGKAM)? GetDataByNBO(string sNbo)
        {
    
            string sSql = $"select T0.\"Code\", COALESCE(T0.\"U_BrandName\", '') as \"U_BrandName\",Coalesce(T1.\"Code\",'') as \"BrandID\",COALESCE(T1.\"U_MkSeg1Name\", '') as \"U_MkSeg1Name\",Coalesce(T1.\"U_Mkseg1\",'') as \"MKSeg1ID\", COALESCE(T1.\"U_OEM\",'') as \"OEM\", COALESCE(T1.\"U_GKAM\",'') as \"GKAM\", Case When Coalesce(T0.\"U_NickName\",'') = '' then Case When Coalesce(T0.\"U_BrandName\",'') = '' then coalesce(T0.\"U_Program\",'') else concat ( T0.\"U_BrandName\",' - ',T0.\"U_Program\") end else concat(T0.\"U_NickName\", ' - ', T0.\"U_BrandName\",' - ',T0.\"U_Program\") end as \"OEM Program\" from \"@STXIXXNBO\" T0 left join \"@STXIXXBRAND\" T1 on T0.\"U_BrandID\" = T1.\"Code\" WHERE T0.\"Code\" = '{sNbo}' ";

            SAPbobsCOM.Recordset oRs = (SAPbobsCOM.Recordset)Utils.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            oRs.DoQuery(sSql);

            if (!oRs.EoF)
            {
                return (
                    oRs.Fields.Item("U_MkSeg1Name").Value.ToString(),
                    oRs.Fields.Item("MKSeg1ID").Value.ToString(),
                    oRs.Fields.Item("U_BrandName").Value.ToString(),
                    oRs.Fields.Item("BrandID").Value.ToString(),
                    oRs.Fields.Item("OEM").Value.ToString(),
                    oRs.Fields.Item("OEM Program").Value.ToString(),
                    oRs.Fields.Item("GKAM").Value.ToString()
                );
            }

            return null;
        }
    }
}