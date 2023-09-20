using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;

namespace STXGen2
{
    internal class DBStructure
    {
        private class UDFDetails
        {
            public string AliasID { get; set; }
            public string FieldDescription { get; set; }
            public BoFieldTypes FieldType { get; set; }
            public int FieldSize { get; set; }
            public BoFldSubTypes EditType { get; set; }
            public string RTable { get; set; }
            public string FieldID { get; set; }
            // Add more properties as needed for other attributes of the UDF.
        }

        private class UDFVVDetails
        {
            public string FieldID { get; set; }
            public string FieldValue { get; set; }
            public string FieldDescription { get; set; }
        }

        public class UDOInfo
        {
            public string UDOName { get; set; }
            public SAPbobsCOM.BoUDOObjType UDOType { get; set; }
        }

        private static readonly Dictionary<string, UDOInfo> UDTtoUDOMapping = new Dictionary<string, UDOInfo>
        {
            { "STXIXXBRAND", new UDOInfo { UDOName = "STXBRAND", UDOType = SAPbobsCOM.BoUDOObjType.boud_MasterData }},
            { "STXIXXMKSEG1", new UDOInfo { UDOName = "MKSEG1", UDOType = SAPbobsCOM.BoUDOObjType.boud_MasterData }},
            { "STXIXXMKSEG2", new UDOInfo { UDOName = "MKSEG2", UDOType = SAPbobsCOM.BoUDOObjType.boud_MasterData }},
            { "STXIXXNBO", new UDOInfo { UDOName = "STXNBOID", UDOType = SAPbobsCOM.BoUDOObjType.boud_MasterData }},
            { "STXOPERATIONS", new UDOInfo { UDOName = "STXOPERATIONS", UDOType = SAPbobsCOM.BoUDOObjType.boud_MasterData }},
            { "STXIXXPRODFAM", new UDOInfo { UDOName = "STXPRFAM", UDOType = SAPbobsCOM.BoUDOObjType.boud_MasterData }},
            { "STXIXXTECHNOLOGIES", new UDOInfo { UDOName = "STXTECH", UDOType = SAPbobsCOM.BoUDOObjType.boud_MasterData }},
            { "STXQC19", new UDOInfo { UDOName = "STXQC19", UDOType = SAPbobsCOM.BoUDOObjType.boud_Document }},
            { "STXSETPTEXTURES", new UDOInfo { UDOName = "SETPTextures", UDOType = SAPbobsCOM.BoUDOObjType.boud_MasterData }}
        };


        private static BoFieldTypes ConvertToBoFieldType(string typeValue)
        {
            switch (typeValue)
            {
                case "A":
                    return BoFieldTypes.db_Alpha;
                case "N":
                    return BoFieldTypes.db_Numeric;
                case "D":
                    return BoFieldTypes.db_Date;
                case "M":
                    return BoFieldTypes.db_Memo;
                case "B":
                    return BoFieldTypes.db_Float;

                default:
                    throw new ArgumentException($"Unsupported field type value: {typeValue}");
            }
        }

        private static BoFldSubTypes ConvertToBoFieldSubType(string subTypeValue)
        {

            switch (subTypeValue)
            {
                case "?":
                    return BoFldSubTypes.st_Address;
                case "#":
                    return BoFldSubTypes.st_Phone;
                case "T":
                    return BoFldSubTypes.st_Time;
                case "R":
                    return BoFldSubTypes.st_Rate;
                case "S":
                    return BoFldSubTypes.st_Sum;
                case "P":
                    return BoFldSubTypes.st_Price;
                case "Q":
                    return BoFldSubTypes.st_Quantity;
                case "%":
                    return BoFldSubTypes.st_Percentage;
                case "M":
                    return BoFldSubTypes.st_Measurement;
                case "B":
                    return BoFldSubTypes.st_Link;
                case "I":
                    return BoFldSubTypes.st_Image;
                case "C":
                    return BoFldSubTypes.st_Checkbox;
                default:
                    return BoFldSubTypes.st_None;
            }
        }

        private static List<UDFDetails> GetUDFsForUDT(string udtName)
        {

            List<UDFDetails> udfs = new List<UDFDetails>();
            Recordset rs = (Recordset)Utils.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            try
            {
                rs.DoQuery($"SELECT * FROM CUFD WHERE TableID = '@{udtName}'");
                while (!rs.EoF)
                {

                    UDFDetails detail = new UDFDetails
                    {
                        AliasID = rs.Fields.Item("AliasID").Value.ToString(),
                        FieldDescription = rs.Fields.Item("Descr").Value.ToString(),
                        FieldType = ConvertToBoFieldType(rs.Fields.Item("TypeID").Value.ToString()),
                        FieldSize = Convert.ToInt32(rs.Fields.Item("EditSize").Value),
                        EditType = ConvertToBoFieldSubType(rs.Fields.Item("EditType").Value.ToString()),
                        RTable = rs.Fields.Item("RTable").Value.ToString(),
                        FieldID = rs.Fields.Item("FieldID").Value.ToString(),
                        // Assign other properties as needed.
                    };
                    udfs.Add(detail);
                    rs.MoveNext();
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                return udfs;
            }
            finally
            {
                if (rs != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                    rs = null;
                    GC.Collect();
                }
            }
        }
        internal static void VerifyTables()
        {
            GC.Collect();
            CreateUDTIfNotExists("STXIXXBRAND", "Intrexx Brands", BoUTBTableType.bott_MasterData);
            CreateUDTIfNotExists("STXIXXMKSEG1", "Intrexx Mkt Seg. 1", BoUTBTableType.bott_MasterData);
            CreateUDTIfNotExists("STXIXXMKSEG2", "Intrexx Mkt Seg. 2", BoUTBTableType.bott_MasterData);
            CreateUDTIfNotExists("STXIXXNBO", "Intrexx NBO", BoUTBTableType.bott_MasterData);
            CreateUDTIfNotExists("STXIXXPRODFAM", "Intrexx Prod Family", BoUTBTableType.bott_MasterData);
            CreateUDTIfNotExists("STXIXXTECHNOLOGIES", "Intrexx Technologies", BoUTBTableType.bott_MasterData);
            CreateUDTIfNotExists("STXOPERATIONS", "Productions Operations", BoUTBTableType.bott_MasterData);

            CreateUDTIfNotExists("STXSETPTEXTURES", "SETP Texture", BoUTBTableType.bott_MasterData);
            CreateUDTIfNotExists("STXSETPTEXTURETASKS", "SETP Texture Tasks", BoUTBTableType.bott_NoObject);

            CreateUDTIfNotExists("STXQC19", "QC19-Header", BoUTBTableType.bott_Document);
            CreateUDTIfNotExists("STXQC19C", "QC19-Other Costs", BoUTBTableType.bott_DocumentLines);
            CreateUDTIfNotExists("STXQC19O", "QC19-Operations", BoUTBTableType.bott_DocumentLines);
            CreateUDTIfNotExists("STXQC19T", "QC19-Textures", BoUTBTableType.bott_DocumentLines);
            CreateUDTIfNotExists("STXQC19TCLASS", "QC19-Texture Class", BoUTBTableType.bott_NoObject);
        }

        private static bool UDTExists(string udtName, out BoUTBTableType currentType)
        {
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utils.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                rs = (Recordset)Utils.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                rs.DoQuery($"SELECT * FROM OUTB WHERE TableName = '{udtName}'");
                if (rs.EoF && rs.BoF)
                {
                    currentType = BoUTBTableType.bott_NoObject; // Just a default value
                    return false;
                }
                currentType = (BoUTBTableType)Enum.Parse(typeof(BoUTBTableType), rs.Fields.Item("ObjectType").Value.ToString());
                return true;

            }
            finally
            {
                if (rs != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                    rs = null;
                    GC.Collect();
                }
            }
        }

        private static void CreateUDTIfNotExists(string udtName, string udtDesc, BoUTBTableType desiredType)
        {
            BoUTBTableType currentType;
            List<UDFDetails> udfs = new List<UDFDetails>();
            List<UDFVVDetails> udfvvs = new List<UDFVVDetails>();
            List<Dictionary<string, object>> data = new List<Dictionary<string, object>>();

            if (UDTExists(udtName, out currentType))
            {
                if (currentType != desiredType)
                {
                    udfs = GetUDFsForUDT(udtName);

                    udfvvs = GetUDFValidValues(udtName);

                    data = FetchAllUDTData(udtName);

                    UserTablesMD udtMD = (UserTablesMD)Utils.oCompany.GetBusinessObject(BoObjectTypes.oUserTables);
                    if (udtMD.GetByKey(udtName))
                    {
                        int result = udtMD.Remove();
                        if (result != 0)
                        {
                            int errCode;
                            string errMsg;
                            Utils.oCompany.GetLastError(out errCode, out errMsg);
                            Console.WriteLine($"Failed to remove UDT. Error: {errCode} - {errMsg}");
                        }
                        else
                        {
                            Console.WriteLine("UDT removed successfully.");
                        }
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(udtMD);
                    }

                    CreateUDT(udtName, udtDesc, desiredType);

                    // Create UDFs for the UDT based on the desired structure
                    foreach (var udfDetail in udfs)
                    {
                        Dictionary<string, string> UdfvvsDic = udfvvs.Where(x => x.FieldID == udfDetail.FieldID).ToDictionary(x => x.FieldValue, x => x.FieldDescription);

                        if (UdfvvsDic.Count > 0) // Check if there are valid values for this field
                        {
                            AddFieldIfNotExists(udtName, udfDetail.AliasID, udfDetail.FieldDescription, udfDetail.FieldType, udfDetail.FieldSize, udfDetail.RTable, UdfvvsDic);
                        }
                        else
                        {
                            // Add field without valid values
                            AddFieldIfNotExists(udtName, udfDetail.AliasID, udfDetail.FieldDescription, udfDetail.FieldType, udfDetail.FieldSize, udfDetail.RTable);
                        }
                    }

                    UDOInfo udoInfo;
                    if (UDTtoUDOMapping.TryGetValue(udtName, out udoInfo))
                    {
                        RegisterUDO(udoInfo.UDOName, udoInfo.UDOName, udoInfo.UDOType, udtName);
                    }

                    // Restore the data if any
                    if (data != null && data.Count > 0)
                    {
                        PopulateUDOWithData(udtName, data);
                    }
                    //PopulateUDTWithData(udtName, data);
                }
            }
            else
            {
                CreateUDT(udtName, udtDesc, desiredType);
            }
        }

        private static List<UDFVVDetails> GetUDFValidValues(string udtName)
        {
            List<UDFVVDetails> udfvvs = new List<UDFVVDetails>();
            Recordset rs = (Recordset)Utils.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            try
            {
                rs.DoQuery($"SELECT * FROM UFD1 WHERE TableID = '@{udtName}'");
                while (!rs.EoF)
                {

                    UDFVVDetails detail = new UDFVVDetails
                    {
                        FieldID = rs.Fields.Item("FieldID").Value.ToString(),
                        FieldValue = rs.Fields.Item("FldValue").Value.ToString(),
                        FieldDescription = rs.Fields.Item("Descr").Value.ToString(),
                        // Assign other properties as needed.
                    };
                    udfvvs.Add(detail);
                    rs.MoveNext();
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                return udfvvs;
            }
            finally
            {
                if (rs != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                    rs = null;
                    GC.Collect();
                }
            }
        }

        private static void RegisterUDO(string udoCode, string udoName, BoUDOObjType udoType , string udtName)
        {
            UserObjectsMD oUDO = null;

            try
            {
                oUDO = (UserObjectsMD)Utils.oCompany.GetBusinessObject(BoObjectTypes.oUserObjectsMD);

                // Check if UDO already exists
                if (!oUDO.GetByKey(udoCode))
                {
                    oUDO.Code = udoCode;
                    oUDO.Name = udoName;
                    oUDO.TableName = udtName;
                    oUDO.ObjectType = udoType;

                    // Assuming that this UDO uses default form
                    oUDO.CanFind = BoYesNoEnum.tYES;
                    oUDO.CanDelete = BoYesNoEnum.tYES;
                    oUDO.ManageSeries = BoYesNoEnum.tNO;
                    oUDO.CanCancel = BoYesNoEnum.tNO;

                    oUDO.EnableEnhancedForm = SAPbobsCOM.BoYesNoEnum.tNO;
                    oUDO.CanCreateDefaultForm = BoYesNoEnum.tYES;




                    // ------------- Find Form fields ----------------

                    var findFormFields = GetFindFormFields(udtName);
                    foreach (var field in findFormFields)
                    {
                        oUDO.FindColumns.ColumnAlias = field.Key;
                        oUDO.FindColumns.ColumnDescription = field.Value;
                        oUDO.FindColumns.Add();
                    }

                    // ------------- Form fields ----------------

                    var udfList = GetUDFsForUDT(udtName);

                    oUDO.FormColumns.FormColumnAlias = "Code";
                    oUDO.FormColumns.FormColumnDescription = "Code";
                    oUDO.FormColumns.Add();
                    oUDO.FormColumns.FormColumnAlias = "Name";
                    oUDO.FormColumns.FormColumnDescription = "Name";
                    oUDO.FormColumns.Add();
                    foreach (var udf in udfList)
                    {
                        oUDO.FormColumns.FormColumnAlias = "U_" + udf.AliasID;
                        oUDO.FormColumns.FormColumnDescription = udf.FieldDescription;
                        oUDO.FormColumns.Add();
                    }

                    int addResult = oUDO.Add();
                    if (addResult != 0)
                    {
                        int errorCode;
                        string errorMsg;
                        Utils.oCompany.GetLastError(out errorCode, out errorMsg);
                        Console.WriteLine($"Error registering UDO: {errorCode} - {errorMsg}");
                    }
                    else
                    {
                        Console.WriteLine("UDO registered successfully!");
                    }
                }
                else
                {
                    Console.WriteLine("UDO already exists.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception: {ex.Message}");
            }
            finally
            {
                if (oUDO != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUDO);
                    oUDO = null;
                    GC.Collect();
                }
            }
        }


        private static void PopulateUDOWithData(string udoName, List<Dictionary<string, object>> data)
        {
            // Get UDO object
            SAPbobsCOM.CompanyService oCompanyService = Utils.oCompany.GetCompanyService();
            SAPbobsCOM.GeneralService oGeneralService = oCompanyService.GetGeneralService(udoName);
            SAPbobsCOM.GeneralData oGeneralData = (SAPbobsCOM.GeneralData)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
            SAPbobsCOM.GeneralDataParams oGeneralParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
            try
            {
                foreach (var row in data)
                {
                    // Assigning fields to the UDO
                    foreach (var field in row)
                    {
                        oGeneralData.SetProperty(field.Key, field.Value);
                    }

                    // Add the data to the UDO
                    SAPbobsCOM.GeneralDataParams response = oGeneralService.Add(oGeneralData);

                    if (response == null || string.IsNullOrEmpty(response.GetProperty("Code").ToString()))
                    {
                        // Handle error
                        int errCode;
                        string errMsg;
                        Utils.oCompany.GetLastError(out errCode, out errMsg);
                        Console.WriteLine($"Failed to add data to UDO '{udoName}'. Error: {errCode} - {errMsg}");
                    }
                }
            }
            finally
            {

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oGeneralService);
                oGeneralService = null;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCompanyService);
                oCompanyService = null;
                GC.Collect();
            }          
        }

        private static Dictionary<string, string> GetFindFormFields(string udtName)
        {
            // You can add as many UDT names and associated fields as you need
            Dictionary<string, Dictionary<string, string>> udtFindFormFields = new Dictionary<string, Dictionary<string, string>>
            {
                {
                    "STXOPERATIONS", new Dictionary<string, string>
                    {
                        {"Code", "Code"},
                        {"U_STXOPDes", "Operation Description"},
                        {"U_STXOPDesLocal", "Operation Description Local"},
                        {"U_STXTechDes", "Technology Description"},
                    }
                },
                // ... add other UDTs as needed
            };

            if (udtFindFormFields.ContainsKey(udtName))
            {
                return udtFindFormFields[udtName];
            }
            else
            {
                // Return an empty dictionary if the UDT name isn't found
                return new Dictionary<string, string>();
            }
        }

        private static void CreateUDT(string udtName, string udtDesc, BoUTBTableType tableType)
        {
            UserTablesMD udtMD = (UserTablesMD)Utils.oCompany.GetBusinessObject(BoObjectTypes.oUserTables);
            try
            {
                udtMD.TableName = udtName;
                udtMD.TableDescription = udtDesc;
                udtMD.TableType = tableType;

                int result = udtMD.Add();
                if (result != 0)
                {
                    int errCode;
                    string errMsg;
                    Utils.oCompany.GetLastError(out errCode, out errMsg);
                    Console.WriteLine($"Failed to create UDT '{udtName}'. Error: {errCode} - {errMsg}");
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(udtMD);
                udtMD = null;
                GC.Collect();
            }
        }

        private static List<Dictionary<string, object>> FetchAllUDTData(string udtName)
        {
            List<Dictionary<string, object>> data = new List<Dictionary<string, object>>();
            Recordset rs = (Recordset)Utils.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            try
            {
                rs.DoQuery($"SELECT * FROM [@{udtName}]");
                while (!rs.EoF)
                {
                    Dictionary<string, object> row = new Dictionary<string, object>();
                    for (int i = 0; i < rs.Fields.Count; i++)
                    {
                        row[rs.Fields.Item(i).Name] = rs.Fields.Item(i).Value;
                    }
                    data.Add(row);
                    rs.MoveNext();
                }
                return data;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                rs = null;
                GC.Collect();
            }
        }


        internal static void VerifyUDF()
        {
            GC.Collect();
            AddFieldIfNotExists("OQUT", "STXOEMPgm", "Brand Program", SAPbobsCOM.BoFieldTypes.db_Alpha,100);
            AddFieldIfNotExists("OQUT", "STXBrand", "Brand", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            AddFieldIfNotExists("OQUT", "STXOEM", "OEM", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            AddFieldIfNotExists("OQUT", "STXMarSeg", "Market Segment 1", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFieldIfNotExists("OQUT", "STXIndCode", "Market Segment 2", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFieldIfNotExists("OQUT", "STXGKAM", "Global KAM", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFieldIfNotExists("OQUT", "STXNBOID", "NBO", SAPbobsCOM.BoFieldTypes.db_Alpha, 10);
            AddFieldIfNotExists("OQUT", "STXMSEGID1", "Market Segment 1 (ID)", SAPbobsCOM.BoFieldTypes.db_Alpha, 10);
            AddFieldIfNotExists("OQUT", "STXMSEGID2", "Market Segment 2 (ID)", SAPbobsCOM.BoFieldTypes.db_Alpha, 10);
            AddFieldIfNotExists("OQUT", "STXBRANDID", "Brand ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 10);

            AddFieldIfNotExists("OQUT", "STXSONum", "Sales order", SAPbobsCOM.BoFieldTypes.db_Numeric, 10);
            AddFieldIfNotExists("OQUT", "STX_CustCode", "Código de Cliente", SAPbobsCOM.BoFieldTypes.db_Alpha, 15);
            AddFieldIfNotExists("OQUT", "STX_Customer", "Nome de Cliente", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);

            AddFieldIfNotExists("OQUT", "STXToolNum", "Tool Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 20);
            


            Dictionary<string, string> saleTypes = new Dictionary<string, string>
            {
                {"DS", "Direct Sale"},
                {"IS", "Inside Sale"}
            };

            AddFieldIfNotExists("OQUT", "SaleType", "Type of Sale", SAPbobsCOM.BoFieldTypes.db_Alpha, 10,null,saleTypes);
            AddFieldIfNotExists("OQUT", "STXRevision", "Revision", SAPbobsCOM.BoFieldTypes.db_Alpha, 2);

            AddFieldIfNotExists("QUT1", "STXPartNum", "Part Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFieldIfNotExists("QUT1", "STXToolNum", "Tool Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 20);
            AddFieldIfNotExists("QUT1", "STXPartName", "Part Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFieldIfNotExists("QUT1", "STXLeadTime", "Lead Time", SAPbobsCOM.BoFieldTypes.db_Numeric, 3);
            AddFieldIfNotExists("QUT1", "STXQC19ID", "Q.Calc ID", SAPbobsCOM.BoFieldTypes.db_Numeric, 10);


            Dictionary<string, string> woType = new Dictionary<string, string>
            {
                {"Normal Production", "Normal Production"},
                {"Tunning", "Tunning"},
                {"Claims", "Claims"},
                {"Internal R&D", "Internal R&D"},
                {"Intercompany Order", "Intercompany Order"},
                {"Internal(Repair/Maintenance)", "Internal(Repair/Maintenance)"},
                {"Revenue - % of Completion", "Revenue - % of Completion"}
            };
            AddFieldIfNotExists("OWOR", "STXSONum", "Sales order", SAPbobsCOM.BoFieldTypes.db_Numeric, 11);
            AddFieldIfNotExists("OWOR", "STXSOLineNum", "SO Line", SAPbobsCOM.BoFieldTypes.db_Numeric, 10);
            AddFieldIfNotExists("OWOR", "STXQC19ID", "Q. Calc ID", SAPbobsCOM.BoFieldTypes.db_Numeric, 11);
            AddFieldIfNotExists("OWOR", "STXWOType", "WO Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 50,null, woType);
            AddFieldIfNotExists("OWOR", "STXWOBrand", "Brand", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            AddFieldIfNotExists("OWOR", "STXCustName", "Fin. Customer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFieldIfNotExists("OWOR", "STXOEMPgm", "Brand Program", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);

            AddFieldIfNotExists("OWOR", "STXSalesEmployee", "Fin. Customer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            AddFieldIfNotExists("OWOR", "STXLicTradNum", "Tax ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 32);

            AddFieldIfNotExists("WOR1", "Texture", "Texture", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFieldIfNotExists("WOR1", "STXOPCode", "Operation Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 20);
            AddFieldIfNotExists("WOR1", "STXOPDes", "Operation Description", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFieldIfNotExists("WOR1", "STXOPDesLocal", "Operation Description Local", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFieldIfNotExists("WOR1", "QCLineID", "QC Operations line ID", SAPbobsCOM.BoFieldTypes.db_Numeric, 10);


        }




        private static bool UDFExists(string tableName, string fieldName)
        {
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utils.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                rs.DoQuery($"SELECT 1 FROM CUFD WHERE TableID = '{tableName}' AND AliasID = '{fieldName}'");
                return !(rs.EoF && rs.BoF);
            }
            finally
            {
                if (rs != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                    rs = null;
                    GC.Collect();
                }
            }
        }

        private static void AddFieldIfNotExists(string tableName, string fieldName, string fieldDescription, SAPbobsCOM.BoFieldTypes fieldType, int fieldSize,string RTable = null , Dictionary<string, string> validValues = null)
        {
            if (!UDFExists(tableName, fieldName))
            {
                SAPbobsCOM.UserFieldsMD uFieldMDLocal = null;
                try
                {
                    uFieldMDLocal = (SAPbobsCOM.UserFieldsMD)Utils.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                    uFieldMDLocal.TableName = tableName;
                    uFieldMDLocal.Name = fieldName;
                    uFieldMDLocal.Description = fieldDescription;
                    uFieldMDLocal.Type = fieldType;
                    uFieldMDLocal.EditSize = fieldSize;

                    uFieldMDLocal.LinkedTable = RTable;

                    if (validValues != null && validValues.Count > 0)
                    {
                        foreach (var validValue in validValues)
                        {
                            uFieldMDLocal.ValidValues.Value = validValue.Key;
                            uFieldMDLocal.ValidValues.Description = validValue.Value;
                            uFieldMDLocal.ValidValues.Add();
                        }
                    }

                    int lRetCode = uFieldMDLocal.Add();
                    if (lRetCode != 0)
                    {
                        int errCode;
                        string errMsg;
                        Utils.oCompany.GetLastError(out errCode, out errMsg);
                        Console.WriteLine($"Failed to create field '{fieldName}'. Error: {errCode} - {errMsg}");
                    }
                }
                finally
                {
                    if (uFieldMDLocal != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(uFieldMDLocal);
                        uFieldMDLocal = null;
                        GC.Collect();
                    }
                }
            }
        }

    }
}