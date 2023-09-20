using System;
using System.Collections.Generic;

public class TableInfo
{
    public string TableDescription { get; set; }
    public string PrimaryKey { get; set; }
    public int ObjectType { get; set; }

    public TableInfo(string description, string primaryKey, int objectType)
    {
        TableDescription = description;
        PrimaryKey = primaryKey;
        ObjectType = objectType;
    }
}

public class Program
{
    public static void Main()
    {
        Dictionary<string, TableInfo> tables = new Dictionary<string, TableInfo>
        {
            {"OACT", new TableInfo("G/L Accounts", "AcctCode", 1)},
            {"OCRD", new TableInfo("Business Partner", "CardCode", 2)},
            {"ODSC", new TableInfo("Bank Codes", "AbsEntry", 3)},
            // ... Continue adding all the tables
        };

        // Example usage
        Console.WriteLine(tables["OACT"].TableDescription);
    }
}
