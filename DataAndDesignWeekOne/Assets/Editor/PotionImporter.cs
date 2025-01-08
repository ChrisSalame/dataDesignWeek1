using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using System.Collections;
using System.Collections.Generic;
using System.IO.Enumeration;
using UnityEngine;


[CreateAssetMenu(fileName = "PotionImporter.asset", menuName = "Data/Potion Importer")]


public class PotionImporter : ScriptableObject
{
    [Tooltip("Path from assets to the excel file to read. (Use forward slashes '/')")]
    public string excelFilePath = "Editor/PotionCrafting.xlsx";

    [Tooltip("Asset to store the loaded recipes.")]
    public RecipeCollection recipes;

    [ContextMenu("Import")]
    public void import()
    {
        var excel = new ExcelImporter(excelFilePath);

        var items = DataHelper.GetAllAssetsOfType<InventoryItem>();


        //This is what imports the ingredients and the potions into the file folders
        ImportItems("Ingredients", excel, items);
        ImportItems("Potions", excel, items);
        ImportRecipes(excel, items);

    }

    void ImportItems(string catagory, ExcelImporter excel, Dictionary<string, InventoryItem> items)
    {
        if (!excel.TryGetTable(catagory, out var table))
        {
            Debug.LogError($"Could not find 'Ingredients' table.");
            return;
        }

        for (int row = 1; row <= table.RowCount; row++)
        {
            string name = table.GetValue<string>(row, "Name");
            if (string.IsNullOrWhiteSpace(name)) continue; //Skip the blank row
            var item = DataHelper.GetOrCreateAsset(name, items, catagory);

            item.cost = table.GetValue<int>(row, "Cost");

            if (string.IsNullOrWhiteSpace(item.displayName))
                item.displayName = name;

            if (table.TryGetEnum<Rarity>(row, "Rarity", out var rarity))
                item.rarity = rarity;

            item.cost = table.GetValue<int>(row, "Cost");

            item.itemType = table.GetValue<string>(row, "itemType");

        }


    }

    void ImportRecipes(ExcelImporter excel, Dictionary<string, InventoryItem> items)
    {
        if (recipes == null)
        {
            Debug.LogError("No recipeCollection provided");
            return;
        }

        if (!excel.TryGetTable("Recipes", out var table))
        {
            Debug.LogError($"Could not find Recipe table.");
            return;
        }

        DataHelper.MarkChangesForSaving(recipes);
        recipes.Clear();

        for (int row = 1; row <= table.RowCount; row++)
        {
            recipes.TryAddRecipe(
                items,
                table.GetValue<string>(row, "Potion"),
                table.GetValue<string>(row, "Item 1"),
                table.GetValue<string>(row, "Item 2"),
                table.GetValue<string>(row, "Item 3")
                );

        }

    }

}
