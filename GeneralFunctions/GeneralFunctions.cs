using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TabularEditor.TOMWrapper;
using TabularEditor.Scripting;
using System.Reflection.Emit;
using Microsoft.VisualBasic;

namespace GeneralFunctions
{

    //copy from the following line up to ****** and remove the // before the closing bracket
    //after the class declaration add all the #r and using statements necessary for the custom class code to run in Tabular Editor
    //these directives will be combined with the ones from the macro when using the CopyMacro script

    public static class Fx
    {
        //#r "Microsoft.VisualBasic"
        //using Microsoft.VisualBasic;


        //in TE2 (at least up to 2.17.2) any method that accesses or modifies the model needs a reference to the model 
        //the following is an example method where you can build extra logic
        public static Table CreateCalcTable(Model model, string tableName, string tableExpression) 
        { 
            return model.AddCalculatedTable(name:tableName,expression:tableExpression);
        }

        public static Table SelectTableExt(Model model, string possibleName = null, string annotationName = null, string annotationValue = null, 
            Func<Table,bool>  lambdaExpression = null, string label = "Select Table", bool skipDialogIfSingleMatch = true, bool showOnlyMatchingTables = true,
            IEnumerable<Table> candidateTables = null, bool showErrorIfNoTablesFound = false, string errorMessage = "No tables found", bool selectFirst = false,
            bool showErrorIfNoSelection = true, string noSelectionErrorMessage = "No table was selected", bool excludeCalcGroups = false,bool returnNullIfNoTablesFound = false)
        {

            Table table = null as Table;

            //Output("lambda expression is null = " + lambdaExpression == null);
            //Output(annotationName + " " + annotationValue);

            if (lambdaExpression == null)
            {
                if (possibleName != null) { 
                    lambdaExpression = (t) => t.Name == possibleName;
                } else if(annotationName!= null && annotationValue != null)
                {
                    lambdaExpression = (t) => t.GetAnnotation(annotationName) == annotationValue;
                }
                else
                {
                    lambdaExpression = (t) => true; //no filtering
                }
            }


            //use candidateTables if passed as argument
            IEnumerable<Table> tables = null as IEnumerable<Table>;

            if(candidateTables != null)
            {
                tables = candidateTables;
            }
            else
            {
                tables = model.Tables;
            }

            //Output("Step 10");
            //Output(tables);

            if(lambdaExpression != null)
            {
                tables = tables.Where(lambdaExpression);
            }

            //Output("Step 20");
            //Output(tables);

            if (excludeCalcGroups)
            {
                tables = tables.Where(t => t.ObjectType != ObjectType.CalculationGroupTable);
            }

            //Output("Step 30");
            //Output(tables);

            //none found, let the user choose from all tables
            if (tables.Count() == 0)
            {

                if (returnNullIfNoTablesFound)
                {
                    if (showErrorIfNoTablesFound) Error(errorMessage);
                    Output("No tables found");
                    return table;
                } 
                else
                {
                    Output("returnNullIfNoTablesFound is false");
                    table =  SelectTable(tables: model.Tables, label: label);
                }
                
            }
            else if (tables.Count() == 1 && !skipDialogIfSingleMatch)
            {
                Output("tables.Count() == 1 && !skipDialogIfSingleMatch");
                table = SelectTable(tables: model.Tables, preselect: tables.First(), label: label);
            }
            else if (tables.Count() == 1 && skipDialogIfSingleMatch)
            {
                table = tables.First();
            } 
            else if (tables.Count() > 1) 
                
            {
                if (selectFirst)
                {
                    table = tables.First();
                }
                else if (showOnlyMatchingTables)
                {
                    Output("showOnlyMatchingTables");
                    table = SelectTable(tables: tables, preselect: tables.First(), label: label);
                }
                else
                {
                    Output("else");
                    table = SelectTable(tables: model.Tables, preselect: tables.First(), label: label);
                }
                
            }
            else
            {
                Error(@"Unexpected logic in ""SelectTableExt""");
                return null;
            }

            if(showErrorIfNoSelection && table == null)
            {
                Error(noSelectionErrorMessage);
            }

            return table;

        }


        public static CalculationGroupTable SelectCalculationGroup(Model model, string possibleName = null, string annotationName = null, string annotationValue = null,
            Func<Table, bool> lambdaExpression = null, string label = "Select Table", bool skipDialogIfSingleMatch = true, bool showOnlyMatchingTables = true,
            bool showErrorIfNoTablesFound = true, string errorMessage = "No calculation groups found",bool selectFirst = false, 
            bool showErrorIfNoSelection = true, string noSelectionErrorMessage = "No calculation group was selected", bool returnNullIfNoTablesFound = false)
        {

            CalculationGroupTable calculationGroupTable = null as CalculationGroupTable;
            
            Func<Table, bool> lambda = (x) => x.ObjectType == ObjectType.CalculationGroupTable;
            if (!model.Tables.Any(lambda)) return calculationGroupTable;

            IEnumerable<Table> tables = model.Tables.Where(lambda);

            //Output(tables.Select(x => x.Name));
            //Output(annotationName + " " + annotationValue);


            Table table = Fx.SelectTableExt(
                model:model,
                possibleName:possibleName,
                annotationName:annotationName,
                annotationValue:annotationValue,
                lambdaExpression:lambdaExpression,
                label:label,
                skipDialogIfSingleMatch:skipDialogIfSingleMatch,
                showOnlyMatchingTables:showOnlyMatchingTables,
                showErrorIfNoTablesFound:showErrorIfNoTablesFound,
                errorMessage:errorMessage, 
                selectFirst:selectFirst,
                showErrorIfNoSelection:showErrorIfNoSelection,
                noSelectionErrorMessage:noSelectionErrorMessage, 
                returnNullIfNoTablesFound:returnNullIfNoTablesFound, 
                candidateTables:tables);

            if(table == null) return calculationGroupTable;

            calculationGroupTable = table as CalculationGroupTable;




            return calculationGroupTable;

        }

        public static CalculationGroupTable AddCalculationGroupExt(Model model, out bool calcGroupWasCreated, string defaultName = "New Calculation Group", 
            string annotationName = null, string annotationValue = null, bool createOnlyIfNotFound = true, 
            string prompt = "Name", string Title = "Provide a name for the Calculation Group", bool customCalcGroupName = true)
        {
            
            Func<Table,bool> lambda = null as Func<Table,bool>;
            CalculationGroupTable cg = null as CalculationGroupTable;
            calcGroupWasCreated = false;
            string calcGroupName = String.Empty;

            if (createOnlyIfNotFound)
            {

                if (annotationName == null && annotationValue == null)
                {

                    if (customCalcGroupName)
                    {
                        calcGroupName = Interaction.InputBox(Prompt: "Name", Title: "Provide a name for the Calculation Group");
                    }
                    else
                    {
                        calcGroupName = defaultName;
                    }

                    cg = Fx.SelectCalculationGroup(model: model, possibleName: calcGroupName, showErrorIfNoTablesFound: false, selectFirst: true);

                }
                else
                {
                    //Output("With annotations");
                    cg = Fx.SelectCalculationGroup(model: model, 
                        showErrorIfNoTablesFound: false, 
                        annotationName: annotationName, 
                        annotationValue: annotationValue, 
                        returnNullIfNoTablesFound: true);
                }

                if (cg != null) return cg;
            }
            
            if (calcGroupName == String.Empty)
            {
                if (customCalcGroupName)
                {
                    calcGroupName = Interaction.InputBox(Prompt: "Name", Title: "Provide a name for the Calculation Group");
                }
                else
                {
                    calcGroupName = defaultName;
                }
            }

            cg = model.AddCalculationGroup(name: calcGroupName);

            if (annotationName != null && annotationValue != null)
            {
                cg.SetAnnotation(annotationName,annotationValue);
            }

            calcGroupWasCreated = true;

            return cg;

        }

        public static CalculationItem AddCalculationItemExt(CalculationGroupTable cg, string calcItemName, string valueExpression = "SELECTEDMEASURE()",
            string formatStringExpression = "", bool createOnlyIfNotFound = true, bool rewriteIfFound = false)
        {

            CalculationItem calcItem = null as CalculationItem;

            Func<CalculationItem, bool> lambda = (ci) => ci.Name == calcItemName;

            if(createOnlyIfNotFound)
            {
                if (cg.CalculationItems.Any(lambda))
                {

                    calcItem = cg.CalculationItems.Where(lambda).FirstOrDefault();

                    if (!rewriteIfFound)
                    {
                        return calcItem;
                    }
                }
            }


            if(calcItem == null)
            {
                calcItem = cg.AddCalculationItem(name: calcItemName, expression: valueExpression);
            }
            else 
            {
                //rewrite the found calcItem
                calcItem.Expression = valueExpression;
            }

            if(formatStringExpression != String.Empty)
            {
                calcItem.FormatStringExpression = formatStringExpression;
            }
            
            return calcItem;
                
        }

        
        


        //add other methods always as "public static" followed by the data type they will return or void if they do not return anything.



        //}

        //******************
        //do not copy from this line below, and remove the // before the closing bracket above to close the class definition
        //do not change the end mark symbol as it is used as is by the copy macro script.


        //Model and Selected cannot be accessed directly. Always pass a reference to the requited objects. 
        //static readonly Model Model;
        //static readonly TabularEditor.UI.UITreeSelection Selected;


        //These functions replicate the ScriptHelper functions so that they can be
        //used inside the script without the ScriptHelper prefix which cannot be used inside tabular editor
        //the list is not complete and does not include all overloads, complete as necessary. 
        public static void Error(string message, int lineNumber = -1, bool suppressHeader = false)
        {
            ScriptHelper.Error(message: message, lineNumber: lineNumber, suppressHeader: suppressHeader);
        }

        public static void Info(string message, int lineNumber = -1)
        {
            ScriptHelper.Info(message: message, lineNumber: lineNumber);
        }

        public static Table SelectTable(IEnumerable<Table> tables, Table preselect = null, string label = "Select Table")
        {
            return ScriptHelper.SelectTable(tables: tables, preselect: preselect, label: label);
        }

        public static void Output(object value, int lineNumber = -1)
        {
            ScriptHelper.Output(value: value, lineNumber: lineNumber);
        }
    }
}
