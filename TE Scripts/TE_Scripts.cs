using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TabularEditor.TOMWrapper;
using TabularEditor.Scripting;
using Microsoft.VisualBasic;
using GeneralFunctions;
using System.Windows.Forms;
using System.Collections;
using System.IO;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using BenchmarkDotNet.Disassemblers;

namespace TE_Scripting
{
    public class TE_Scripts
    {

        void timeIntelligenceCalcGroupCreation()
        {
            //#r "Microsoft.VisualBasic"
            //using Microsoft.VisualBasic;
            //
            // CHANGELOG:
            // '2021-05-01 / B.Agullo / 
            // '2021-05-17 / B.Agullo / added affected measure table
            // '2021-06-19 / B.Agullo / data label measures
            // '2021-07-10 / B.Agullo / added flag expression to avoid breaking already special format strings
            // '2021-09-23 / B.Agullo / added code to prompt for parameters (code credit to Daniel Otykier) 
            // '2021-09-27 / B.Agullo / added code for general name 
            // '2022-10-11 / B.Agullo / added MMT and MWT calc item groups
            // '2023-01-24 / B.Agullo / added Date Range Measure and completed dynamic label for existing items
            //
            // by Bernat Agulló
            // twitter: @AgulloBernat
            // www.esbrina-ba.com/blog
            //
            // REFERENCE: 
            // Check out https://www.esbrina-ba.com/time-intelligence-the-smart-way/ where this script is introduced
            // 
            // FEATURED: 
            // this script featured in GuyInACube https://youtu.be/_j0iTUo2HT0
            //
            // THANKS:
            // shout out to Johnny Winter for the base script and SQLBI for daxpatterns.com

            //select the measures that you want to be affected by the calculation group
            //before running the script. 
            //measure names can also be included in the following array (no need to select them) 
            string[] preSelectedMeasures = { }; //include measure names in double quotes, like: {"Profit","Total Cost"};

            //AT LEAST ONE MEASURE HAS TO BE AFFECTED!, 
            //either by selecting it or typing its name in the preSelectedMeasures Variable



            //
            // ----- do not modify script below this line -----
            //


            string affectedMeasures = "{";

            int i = 0;

            for (i = 0; i < preSelectedMeasures.GetLength(0); i++)
            {

                if (affectedMeasures == "{")
                {
                    affectedMeasures = affectedMeasures + "\"" + preSelectedMeasures[i] + "\"";
                }
                else
                {
                    affectedMeasures = affectedMeasures + ",\"" + preSelectedMeasures[i] + "\"";
                };

            };


            if (Selected.Measures.Count != 0)
            {

                foreach (var m in Selected.Measures)
                {
                    if (affectedMeasures == "{")
                    {
                        affectedMeasures = affectedMeasures + "\"" + m.Name + "\"";
                    }
                    else
                    {
                        affectedMeasures = affectedMeasures + ",\"" + m.Name + "\"";
                    };
                };
            };

            //check that by either method at least one measure is affected
            if (affectedMeasures == "{")
            {
                Error("No measures affected by calc group");
                return;
            };

            string calcGroupName = String.Empty;
            string columnName = String.Empty;

            if (Model.CalculationGroups.Any(cg => cg.GetAnnotation("@AgulloBernat") == "Time Intel Calc Group"))
            {
                calcGroupName = Model.CalculationGroups.Where(cg => cg.GetAnnotation("@AgulloBernat") == "Time Intel Calc Group").First().Name;

            }
            else
            {
                calcGroupName = Interaction.InputBox("Provide a name for your Calc Group", "Calc Group Name", "Time Intelligence", 740, 400);
            };

            if (calcGroupName == String.Empty) return;


            if (Model.CalculationGroups.Any(cg => cg.GetAnnotation("@AgulloBernat") == "Time Intel Calc Group"))
            {
                columnName = Model.Tables.Where(cg => cg.GetAnnotation("@AgulloBernat") == "Time Intel Calc Group").First().Columns.First().Name;

            }
            else
            {
                columnName = Interaction.InputBox("Provide a name for your Calc Group Column", "Calc Group Column Name", calcGroupName, 740, 400);
            };

            if (columnName == String.Empty) return;

            string affectedMeasuresTableName = String.Empty;

            if (Model.Tables.Any(t => t.GetAnnotation("@AgulloBernat") == "Time Intel Affected Measures Table"))
            {
                affectedMeasuresTableName = Model.Tables.Where(t => t.GetAnnotation("@AgulloBernat") == "Time Intel Affected Measures Table").First().Name;

            }
            else
            {
                affectedMeasuresTableName = Interaction.InputBox("Provide a name for affected measures table", "Affected Measures Table Name", calcGroupName + " Affected Measures", 740, 400);

            };

            if (affectedMeasuresTableName ==String.Empty) return;


            string affectedMeasuresColumnName = String.Empty;

            if (Model.Tables.Any(t => t.GetAnnotation("@AgulloBernat") == "Time Intel Affected Measures Table"))
            {
                affectedMeasuresColumnName = Model.Tables.Where(t => t.GetAnnotation("@AgulloBernat") == "Time Intel Affected Measures Table").First().Columns.First().Name;

            }
            else
            {
                affectedMeasuresColumnName = Interaction.InputBox("Provide a name for affected measures column", "Affected Measures Table Column Name", "Measure", 740, 400);

            };

            if (affectedMeasuresColumnName == String.Empty) return;
            //string affectedMeasuresColumnName = "Measure"; 

            string labelAsValueMeasureName = "Label as Value Measure";
            string labelAsFormatStringMeasureName = "Label as format string";


            // '2021-09-24 / B.Agullo / model object selection prompts! 
            var factTable = SelectTable(label: "Select your fact table");
            if (factTable == null) return;

            var factTableDateColumn = SelectColumn(factTable.Columns, label: "Select the main date column");
            if (factTableDateColumn == null) return;

            Table dateTableCandidate = null;

            if (Model.Tables.Any
                (x => x.GetAnnotation("@AgulloBernat") == "Time Intel Date Table"
                    || x.Name == "Date"
                    || x.Name == "Calendar"))
            {
                dateTableCandidate = Model.Tables.Where
                    (x => x.GetAnnotation("@AgulloBernat") == "Time Intel Date Table"
                        || x.Name == "Date"
                        || x.Name == "Calendar").First();

            };

            var dateTable =
                SelectTable(
                    label: "Select your date table",
                    preselect: dateTableCandidate);

            if (dateTable == null)
            {
                Error("You just aborted the script");
                return;
            }
            else
            {
                dateTable.SetAnnotation("@AgulloBernat", "Time Intel Date Table");
            };


            Column dateTableDateColumnCandidate = null;

            if (dateTable.Columns.Any
                        (x => x.GetAnnotation("@AgulloBernat") == "Time Intel Date Table Date Column" || x.Name == "Date"))
            {
                dateTableDateColumnCandidate = dateTable.Columns.Where
                    (x => x.GetAnnotation("@AgulloBernat") == "Time Intel Date Table Date Column" || x.Name == "Date").First();
            };

            var dateTableDateColumn =
                SelectColumn(
                    dateTable.Columns,
                    label: "Select the date column",
                    preselect: dateTableDateColumnCandidate);

            if (dateTableDateColumn == null)
            {
                Error("You just aborted the script");
                return;
            }
            else
            {
                dateTableDateColumn.SetAnnotation("@AgulloBernat", "Time Intel Date Table Date Column");
            };

            Column dateTableYearColumnCandidate = null;
            if (dateTable.Columns.Any(x => x.GetAnnotation("@AgulloBernat") == "Time Intel Date Table Year Column" || x.Name == "Year"))
            {
                dateTable.Columns.Where
                    (x => x.GetAnnotation("@AgulloBernat") == "Time Intel Date Table Year Column" || x.Name == "Year").First();
            };

            var dateTableYearColumn =
                SelectColumn(
                    dateTable.Columns,
                    label: "Select the year column",
                    preselect: dateTableYearColumnCandidate);

            if (dateTableYearColumn == null)
            {
                Error("You just abourted the script");
                return;
            }
            else
            {
                dateTableYearColumn.SetAnnotation("@AgulloBernat", "Time Intel Date Table Year Column");
            };


            //these names are for internal use only, so no need to be super-fancy, better stick to datpatterns.com model
            string ShowValueForDatesMeasureName = "ShowValueForDates";
            string dateWithSalesColumnName = "DateWith" + factTable.Name;

            // '2021-09-24 / B.Agullo / I put the names back to variables so I don't have to tough the script
            string factTableName = factTable.Name;
            string factTableDateColumnName = factTableDateColumn.Name;
            string dateTableName = dateTable.Name;
            string dateTableDateColumnName = dateTableDateColumn.Name;
            string dateTableYearColumnName = dateTableYearColumn.Name;

            // '2021-09-24 / B.Agullo / this is for internal use only so better leave it as is 
            string flagExpression = "UNICHAR( 8204 )";

            string calcItemProtection = "<CODE>"; //default value if user has selected no measures
            string calcItemFormatProtection = "<CODE>"; //default value if user has selected no measures

            // check if there's already an affected measure table
            if (Model.Tables.Any(t => t.GetAnnotation("@AgulloBernat") == "Time Intel Affected Measures Table"))
            {
                //modifying an existing calculated table is not risk-free
                Info("Make sure to include measure names to the table " + affectedMeasuresTableName);
            }
            else
            {
                // create calculated table containing all names of affected measures
                // this is why you need to enable 
                if (affectedMeasures != "{")
                {

                    affectedMeasures = affectedMeasures + "}";

                    string affectedMeasureTableExpression =
                        "SELECTCOLUMNS(" + affectedMeasures + ",\"" + affectedMeasuresColumnName + "\",[Value])";

                    var affectedMeasureTable =
                        Model.AddCalculatedTable(affectedMeasuresTableName, affectedMeasureTableExpression);

                    affectedMeasureTable.FormatDax();
                    affectedMeasureTable.Description =
                        "Measures affected by " + calcGroupName + " calculation group.";

                    affectedMeasureTable.SetAnnotation("@AgulloBernat", "Time Intel Affected Measures Table");

                    // this causes error
                    // affectedMeasureTable.Columns[affectedMeasuresColumnName].SetAnnotation("@AgulloBernat","Time Intel Affected Measures Table Column");

                    affectedMeasureTable.IsHidden = true;

                };
            };

            //if there where selected or preselected measures, prepare protection code for expresion and formatstring
            string affectedMeasuresValues = "VALUES('" + affectedMeasuresTableName + "'[" + affectedMeasuresColumnName + "])";

            calcItemProtection =
                "SWITCH(" +
                "   TRUE()," +
                "   SELECTEDMEASURENAME() IN " + affectedMeasuresValues + "," +
                "   <CODE> ," +
                "   ISSELECTEDMEASURE([" + labelAsValueMeasureName + "])," +
                "   <LABELCODE> ," +
                "   SELECTEDMEASURE() " +
                ")";


            calcItemFormatProtection =
                "SWITCH(" +
                "   TRUE() ," +
                "   SELECTEDMEASURENAME() IN " + affectedMeasuresValues + "," +
                "   <CODE> ," +
                "   ISSELECTEDMEASURE([" + labelAsFormatStringMeasureName + "])," +
                "   <LABELCODEFORMATSTRING> ," +
                "   SELECTEDMEASUREFORMATSTRING() " +
                ")";


            string dateColumnWithTable = "'" + dateTableName + "'[" + dateTableDateColumnName + "]";
            string yearColumnWithTable = "'" + dateTableName + "'[" + dateTableYearColumnName + "]";
            string factDateColumnWithTable = "'" + factTableName + "'[" + factTableDateColumnName + "]";
            string dateWithSalesWithTable = "'" + dateTableName + "'[" + dateWithSalesColumnName + "]";
            string calcGroupColumnWithTable = "'" + calcGroupName + "'[" + columnName + "]";

            //check to see if a table with this name already exists
            //if it doesnt exist, create a calculation group with this name
            if (!Model.Tables.Contains(calcGroupName))
            {
                var cg = Model.AddCalculationGroup(calcGroupName);
                cg.Description = "Calculation group for time intelligence. Availability of data is taken from " + factTableName + ".";
                cg.SetAnnotation("@AgulloBernat", "Time Intel Calc Group");
            };

            //set variable for the calc group
            Table calcGroup = Model.Tables[calcGroupName];

            //if table already exists, make sure it is a Calculation Group type
            if (calcGroup.SourceType.ToString() != "CalculationGroup")
            {
                Error("Table exists in Model but is not a Calculation Group. Rename the existing table or choose an alternative name for your Calculation Group.");
                return;
            };

            //adds the two measures that will be used for label as value, label as format string 
            var labelAsValueMeasure = calcGroup.AddMeasure(labelAsValueMeasureName, "");
            labelAsValueMeasure.Description = "Use this measure to show the year evaluated in tables";

            var labelAsFormatStringMeasure = calcGroup.AddMeasure(labelAsFormatStringMeasureName, "0");
            labelAsFormatStringMeasure.Description = "Use this measure to show the year evaluated in charts";

            //by default the calc group has a column called Name. If this column is still called Name change this in line with specfied variable
            if (calcGroup.Columns.Contains("Name"))
            {
                calcGroup.Columns["Name"].Name = columnName;

            };

            calcGroup.Columns[columnName].Description = "Select value(s) from this column to apply time intelligence calculations.";
            calcGroup.Columns[columnName].SetAnnotation("@AgulloBernat", "Time Intel Calc Group Column");


            //Only create them if not in place yet (reruns)
            if (!Model.Tables[dateTableName].Columns.Any(C => C.GetAnnotation("@AgulloBernat") == "Date with Data Column"))
            {
                string DateWithSalesCalculatedColumnExpression =
                    dateColumnWithTable + " <= MAX ( " + factDateColumnWithTable + ")";

                Column dateWithDataColumn = dateTable.AddCalculatedColumn(dateWithSalesColumnName, DateWithSalesCalculatedColumnExpression);
                dateWithDataColumn.SetAnnotation("@AgulloBernat", "Date with Data Column");
            };

            if (!Model.Tables[dateTableName].Measures.Any(M => M.Name == ShowValueForDatesMeasureName))
            {
                string ShowValueForDatesMeasureExpression =
                    "VAR LastDateWithData = " +
                    "    CALCULATE ( " +
                    "        MAX (  " + factDateColumnWithTable + " ), " +
                    "        REMOVEFILTERS () " +
                    "    )" +
                    "VAR FirstDateVisible = " +
                    "    MIN ( " + dateColumnWithTable + " ) " +
                    "VAR Result = " +
                    "    FirstDateVisible <= LastDateWithData " +
                    "RETURN " +
                    "    Result ";

                var ShowValueForDatesMeasure = dateTable.AddMeasure(ShowValueForDatesMeasureName, ShowValueForDatesMeasureExpression);

                ShowValueForDatesMeasure.FormatDax();
            };



            //defining expressions and formatstring for each calc item
            string CY =
                "/*CY*/ " +
                "SELECTEDMEASURE()";

            string CYlabel =
                "SELECTEDVALUE(" + yearColumnWithTable + ")";


            string PY =
                "/*PY*/ " +
                "IF (" +
                "    [" + ShowValueForDatesMeasureName + "], " +
                "    CALCULATE ( " +
                "        " + CY + ", " +
                "        CALCULATETABLE ( " +
                "            DATEADD ( " + dateColumnWithTable + " , -1, YEAR ), " +
                "            " + dateWithSalesWithTable + " = TRUE " +
                "        ) " +
                "    ) " +
                ") ";


            string PYlabel =
                "/*PY*/ " +
                "IF (" +
                "    [" + ShowValueForDatesMeasureName + "], " +
                "    CALCULATE ( " +
                "        " + CYlabel + ", " +
                "        CALCULATETABLE ( " +
                "            DATEADD ( " + dateColumnWithTable + " , -1, YEAR ), " +
                "            " + dateWithSalesWithTable + " = TRUE " +
                "        ) " +
                "    ) " +
                ") ";


            string YOY =
                "/*YOY*/ " +
                "VAR ValueCurrentPeriod = " + CY + " " +
                "VAR ValuePreviousPeriod = " + PY + " " +
                "VAR Result = " +
                "IF ( " +
                "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " +
                "     ValueCurrentPeriod - ValuePreviousPeriod" +
                " ) " +
                "RETURN " +
                "   Result ";

            string YOYlabel =
                "/*YOY*/ " +
                "VAR ValueCurrentPeriod = " + CYlabel + " " +
                "VAR ValuePreviousPeriod = " + PYlabel + " " +
                "VAR Result = " +
                "IF ( " +
                "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " +
                "     ValueCurrentPeriod & \" vs \" & ValuePreviousPeriod" +
                " ) " +
                "RETURN " +
                "   Result ";

            string YOYpct =
                "/*YOY%*/ " +
               "VAR ValueCurrentPeriod = " + CY + " " +
                "VAR ValuePreviousPeriod = " + PY + " " +
                "VAR CurrentMinusPreviousPeriod = " +
                "IF ( " +
                "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " +
                "     ValueCurrentPeriod - ValuePreviousPeriod" +
                " ) " +
                "VAR Result = " +
                "DIVIDE ( " +
                "    CurrentMinusPreviousPeriod," +
                "    ValuePreviousPeriod" +
                ") " +
                "RETURN " +
                "  Result";

            string YOYpctLabel =
                "/*YOY%*/ " +
               "VAR ValueCurrentPeriod = " + CYlabel + " " +
                "VAR ValuePreviousPeriod = " + PYlabel + " " +
                "VAR Result = " +
                "IF ( " +
                "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " +
                "     ValueCurrentPeriod & \" vs \" & ValuePreviousPeriod & \" (%)\"" +
                " ) " +
                "RETURN " +
                "  Result";

            string YTD =
                "/*YTD*/" +
                "IF (" +
                "    [" + ShowValueForDatesMeasureName + "]," +
                "    CALCULATE (" +
                "        " + CY + "," +
                "        DATESYTD (" + dateColumnWithTable + " )" +
                "   )" +
                ") ";


            string YTDlabel = CYlabel + "& \" YTD\"";


            string PYTD =
                "/*PYTD*/" +
                "IF ( " +
                "    [" + ShowValueForDatesMeasureName + "], " +
                "   CALCULATE ( " +
                "       " + YTD + "," +
                "    CALCULATETABLE ( " +
                "        DATEADD ( " + dateColumnWithTable + ", -1, YEAR ), " +
                "       " + dateWithSalesWithTable + " = TRUE " +
                "       )" +
                "   )" +
                ") ";

            string PYTDlabel = PYlabel + "& \" YTD\"";


            string YOYTD =
                "/*YOYTD*/" +
                "VAR ValueCurrentPeriod = " + YTD + " " +
                "VAR ValuePreviousPeriod = " + PYTD + " " +
                "VAR Result = " +
                "IF ( " +
                "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " +
                "     ValueCurrentPeriod - ValuePreviousPeriod" +
                " ) " +
                "RETURN " +
                "   Result ";


            string YOYTDlabel =
                "/*YOYTD*/" +
                "VAR ValueCurrentPeriod = " + YTDlabel + " " +
                "VAR ValuePreviousPeriod = " + PYTDlabel + " " +
                "VAR Result = " +
                "IF ( " +
                "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " +
                "     ValueCurrentPeriod & \" vs \" & ValuePreviousPeriod" +
                " ) " +
                "RETURN " +
                "   Result ";



            string YOYTDpct =
                "/*YOYTD%*/" +
                "VAR ValueCurrentPeriod = " + YTD + " " +
                "VAR ValuePreviousPeriod = " + PYTD + " " +
                "VAR CurrentMinusPreviousPeriod = " +
                "IF ( " +
                "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " +
                "     ValueCurrentPeriod - ValuePreviousPeriod" +
                " ) " +
                "VAR Result = " +
                "DIVIDE ( " +
                "    CurrentMinusPreviousPeriod," +
                "    ValuePreviousPeriod" +
                ") " +
                "RETURN " +
                "  Result";


            string YOYTDpctLabel =
                "/*YOY%*/ " +
               "VAR ValueCurrentPeriod = " + YTDlabel + " " +
                "VAR ValuePreviousPeriod = " + PYTDlabel + " " +
                "VAR Result = " +
                "IF ( " +
                "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " +
                "     ValueCurrentPeriod & \" vs \" & ValuePreviousPeriod & \" (%)\"" +
                " ) " +
                "RETURN " +
                "  Result";


            string MAT =
             "        /*TAM*/" +
             "        IF (" +
                "    [" + ShowValueForDatesMeasureName + "], " +
             "            CALCULATE (" +
             "                SELECTEDMEASURE()," +
             "                DATESINPERIOD (" +
             "                    " + dateColumnWithTable + " ," +
             "                    MAX ( " + dateColumnWithTable + "  )," +
             "                    -1," +
             "                    YEAR" +
             "                )" +
             "                " +
             "            )" +
             "        )";


            string MATlabel =
                "        /*TAM*/" +
             "        IF (" +
                "    [" + ShowValueForDatesMeasureName + "], " +
             "            CALCULATE (" +
             "                \"Year ending \" & FORMAT(MAX( 'Date'[Date] ),\"d-MMM-yyyy\",\"en-US\")," +
             "                DATESINPERIOD (" +
             "                    " + dateColumnWithTable + " ," +
             "                    MAX ( " + dateColumnWithTable + "  )," +
             "                    -1," +
             "                    YEAR" +
             "                )" +
             "                " +
             "            )" +
             "        )";

            string MATminus1 =
             "        /*TAM*/" +
             "        IF (" +
             "            [" + ShowValueForDatesMeasureName + "], " +
             "            CALCULATE (" +
             "                SELECTEDMEASURE()," +
             "                DATESINPERIOD (" +
             "                    " + dateColumnWithTable + "," +
             "                    LASTDATE( DATEADD( " + dateColumnWithTable + ", - 1, YEAR ) )," +
             "                    -1," +
             "                    YEAR" +
             "                )" +
             "            )" +
             "        )";

            string MATminus1label = 
                "/*MAT-1*/" +
             "        IF (" +
             "            [" + ShowValueForDatesMeasureName + "], " +
             "            CALCULATE (" +
             "                \"Year ending \" & FORMAT(MAX( 'Date'[Date] ),\"d-MMM-yyyy\",\"en-US\")," +
             "                DATESINPERIOD (" +
             "                    " + dateColumnWithTable + "," +
             "                    LASTDATE( DATEADD( " + dateColumnWithTable + ", - 1, YEAR ) )," +
             "                    -1," +
             "                    YEAR" +
             "                )" +
             "            )" +
             "        )";
            ;

            string MATvsMATminus1 =
             "        /*MAT vs MAT-1*/\r\n" +
             "        VAR MAT = " + MAT + "\r\n" +
             "        VAR MAT_1 =" + MATminus1 + "\r\n" +
             "        RETURN \r\n" +
             "            IF( ISBLANK( MAT ) || ISBLANK( MAT_1 ), BLANK(), MAT - MAT_1 )";

            string MATvsMATminus1label = "/*MAT vs MAT-1*/" +
                
             "        VAR MAT = " + MATlabel + "\r\n" +
             "        VAR MAT_1 =" + MATminus1label + "\r\n" +
             "        RETURN \r\n" +
             "            IF( ISBLANK( MAT ) || ISBLANK( MAT_1 ), BLANK(), MAT & \" vs \" & MAT_1 )";

            string MATvsMATminus1pct =
             "        /*MAT vs MAT-1(%)*/" +
             "        VAR MAT = " + MAT + "\r\n" +
             "        VAR MAT_1 =" + MATminus1 + "\r\n" +
             "        RETURN" +
             "            IF(" +
             "                ISBLANK( MAT ) || ISBLANK( MAT_1 )," +
             "                BLANK()," +
             "                DIVIDE( MAT - MAT_1, MAT_1 )" +
             "            )";

            string MATvsMATminus1pctlabel = "/*MAT vs MAT-1 (%)*/" +
                             "        VAR MAT = " + MATlabel + "\r\n" +
             "        VAR MAT_1 =" + MATminus1label + "\r\n" +
             "        RETURN \r\n" +
             "            IF( ISBLANK( MAT ) || ISBLANK( MAT_1 ), BLANK(), MAT & \" vs \" & MAT_1 & \" (%)\" )"; 

            string MMT = String.Format(
                    @"/*MMT*/
        IF(
            [{0}],
            CALCULATE( SELECTEDMEASURE( ), DATESINPERIOD( {1}, MAX( {1} ), -1, MONTH ) )
        )", ShowValueForDatesMeasureName, dateColumnWithTable);

            string MMTlabel = String.Format(
                    @"/*MMT*/
        IF(
            [{0}],
            CALCULATE( {2}, DATESINPERIOD( {1}, MAX( {1} ), -1, MONTH ) )
        )", ShowValueForDatesMeasureName, dateColumnWithTable, "\"Month ending \" & FORMAT(MAX( 'Date'[Date] ),\"d-MMM-yyyy\",\"en-US\")");

            string MMTminus1 = String.Format(
                    @"/*MMT*/
        IF(
            [{0}],
            CALCULATE( SELECTEDMEASURE( ), DATESINPERIOD( {1}, LASTDATE( DATEADD( {1}, -1, MONTH ) ), -1, MONTH ) )
        )", ShowValueForDatesMeasureName, dateColumnWithTable);

            string MMTminus1label = "/*MMT-1*/" +
                String.Format(
                    @"/*MMT*/
        IF(
            [{0}],
            CALCULATE( {2}, DATESINPERIOD( {1}, LASTDATE( DATEADD( {1}, -1, MONTH ) ), -1, MONTH ) )
        )", ShowValueForDatesMeasureName, dateColumnWithTable, "\"Month ending \" & FORMAT(MAX( 'Date'[Date] ),\"d-MMM-yyyy\",\"en-US\")");

            string MMTvsMMTminus1 =
             "        /*MMT vs MMT-1*/\r\n" +
             "        VAR MMT = " + MMT + "\r\n" +
             "        VAR MMT_1 =" + MMTminus1 + "\r\n" +
             "        RETURN \r\n" +
             "            IF( ISBLANK( MMT ) || ISBLANK( MMT_1 ), BLANK(), MMT - MMT_1 )";

            string MMTvsMMTminus1label =
                "        /*MMT vs MMT-1*/\r\n" +
             "        VAR MMT = " + MMTlabel + "\r\n" +
             "        VAR MMT_1 =" + MMTminus1label + "\r\n" +
             "        RETURN \r\n" +
             "            IF( ISBLANK( MMT ) || ISBLANK( MMT_1 ), BLANK(), MMT & \" vs \" & MMT_1 )"; 

            string MMTvsMMTminus1pct =
             "        /*MMT vs MMT-1(%)*/" +
             "        VAR MMT = " + MMT + "\r\n" +
             "        VAR MMT_1 =" + MMTminus1 + "\r\n" +
             "        RETURN" +
             "            IF(" +
             "                ISBLANK( MMT ) || ISBLANK( MMT_1 )," +
             "                BLANK()," +
             "                DIVIDE( MMT - MMT_1, MMT_1 )" +
             "            )";

            string MMTvsMMTminus1pctlabel =
                "        /*MMT vs MMT-1(%)*/" +
             "        VAR MMT = " + MMTlabel + "\r\n" +
             "        VAR MMT_1 =" + MMTminus1label + "\r\n" +
             "        RETURN" +
             "            IF( ISBLANK( MMT ) || ISBLANK( MMT_1 ), BLANK(), MMT & \" vs \" & MMT_1  & \" (%)\")";



            string MWT = String.Format(
                    @"/*MWT*/
        IF(
            [{0}],
            CALCULATE( SELECTEDMEASURE( ), DATESINPERIOD( {1}, MAX( {1} ), -7, DAY ) )
        )", ShowValueForDatesMeasureName, dateColumnWithTable);

            string MWTlabel = "/*MWT*/" +
                
                String.Format(
                    @"/*MWT*/
        IF(
            [{0}],
            CALCULATE( {2}, DATESINPERIOD( {1}, MAX( {1} ), -7, DAY ) )
        )", ShowValueForDatesMeasureName, dateColumnWithTable, "\"Week ending \" & FORMAT(MAX( 'Date'[Date] ),\"d-MMM-yyyy\",\"en-US\")"); ;

            string MWTminus1 = String.Format(
                    @"/*MWT*/
        IF(
            [{0}],
            CALCULATE( SELECTEDMEASURE( ), DATESINPERIOD( {1}, LASTDATE( DATEADD( {1}, -7, DAY ) ), -7, DAY ) )
        )", ShowValueForDatesMeasureName, dateColumnWithTable);

            string MWTminus1label = "/*MWT-1*/" +
                String.Format(
                    @"/*MWT*/
        IF(
            [{0}],
            CALCULATE( {2}, DATESINPERIOD( {1}, LASTDATE( DATEADD( {1}, -7, DAY ) ), -7, DAY ) )
        )", ShowValueForDatesMeasureName, dateColumnWithTable, "\"Week ending \" & FORMAT(MAX( 'Date'[Date] ),\"d-MMM-yyyy\",\"en-US\")");

            string MWTvsMWTminus1 =
             "        /*MWT vs MWT-1*/\r\n" +
             "        VAR MWT = " + MWT + "\r\n" +
             "        VAR MWT_1 =" + MWTminus1 + "\r\n" +
             "        RETURN \r\n" +
             "            IF( ISBLANK( MWT ) || ISBLANK( MWT_1 ), BLANK(), MWT - MWT_1 )";

            string MWTvsMWTminus1label = 
                "        /*MWT vs MWT-1*/\r\n" +
             "        VAR MWT = " + MWTlabel + "\r\n" +
             "        VAR MWT_1 =" + MWTminus1label + "\r\n" +
             "        RETURN \r\n" +
             "            IF( ISBLANK( MWT ) || ISBLANK( MWT_1 ), BLANK(), MWT & \" vs \" & MWT_1 )"; 

            string MWTvsMWTminus1pct =
             "        /*MWT vs MWT-1(%)*/" +
             "        VAR MWT = " + MWT + "\r\n" +
             "        VAR MWT_1 =" + MWTminus1 + "\r\n" +
             "        RETURN" +
             "            IF(" +
             "                ISBLANK( MWT ) || ISBLANK( MWT_1 )," +
             "                BLANK()," +
             "                DIVIDE( MWT - MWT_1, MWT_1 )" +
             "            )";

            string MWTvsMWTminus1pctlabel = 
                "/*MWT vs MWT-1 (%)*/" +
             "        VAR MWT = " + MWTlabel + "\r\n" +
             "        VAR MWT_1 =" + MWTminus1label + "\r\n" +
             "        RETURN \r\n" +
             "            IF( ISBLANK( MWT ) || ISBLANK( MWT_1 ), BLANK(), MWT & \" vs \" & MWT_1 & \" (%)\")";



            string defFormatString = "SELECTEDMEASUREFORMATSTRING()";

            //if the flag expression is already present in the format string, do not change it, otherwise apply % format. 
            string pctFormatString =
            "IF(" +
            "\n	FIND( " + flagExpression + ", SELECTEDMEASUREFORMATSTRING(), 1, - 1 ) <> -1," +
            "\n	SELECTEDMEASUREFORMATSTRING()," +
            "\n	\"#,##0.# %\"" +
            "\n)";


            //the order in the array also determines the ordinal position of the item    
            string[,] calcItems =
                {
        {"CY",      CY,         defFormatString,    "Current year",             CYlabel},
        {"PY",      PY,         defFormatString,    "Previous year",            PYlabel},
        {"YOY",     YOY,        defFormatString,    "Year-over-year",           YOYlabel},
        {"YOY%",    YOYpct,     pctFormatString,    "Year-over-year%",          YOYpctLabel},
        {"YTD",     YTD,        defFormatString,    "Year-to-date",             YTDlabel},
        {"PYTD",    PYTD,       defFormatString,    "Previous year-to-date",    PYTDlabel},
        {"YOYTD",   YOYTD,      defFormatString,    "Year-over-year-to-date",   YOYTDlabel},
        {"YOYTD%",  YOYTDpct,   pctFormatString,    "Year-over-year-to-date%",  YOYTDpctLabel},
        {"MAT",     MAT,        defFormatString,    "Moving Anual Total",       MATlabel},
        {"MAT-1",   MATminus1,  defFormatString,    "Moving Anual Total -1 year", MATminus1label},
        {"MAT vs MAT-1", MATvsMATminus1, defFormatString, "Moving Anual Total vs Moving Anual Total -1 year", MATvsMATminus1label},
        {"MAT vs MAT-1(%)", MATvsMATminus1pct, pctFormatString, "Moving Anual Total vs Moving Anual Total -1 year (%)", MATvsMATminus1pctlabel},
        {"MMT",     MMT,        defFormatString,    "Moving Monthly Total",       MMTlabel},
        {"MMT-1",   MMTminus1,  defFormatString,    "Moving Monthly Total -1 month", MMTminus1label},
        {"MMT vs MMT-1", MMTvsMMTminus1, defFormatString, "Moving Monthly Total vs Moving Monthly Total -1 month", MMTvsMMTminus1label},
        {"MMT vs MMT-1(%)", MMTvsMMTminus1pct, pctFormatString, "Moving Monthly Total vs Moving Monthly Total -1 month (%)", MMTvsMMTminus1pctlabel},
        {"MWT",     MWT,        defFormatString,    "Moving Weekly Total",       MWTlabel},
        {"MWT-1",   MWTminus1,  defFormatString,    "Moving Weekly Total -1 week", MWTminus1label},
        {"MWT vs MWT-1", MWTvsMWTminus1, defFormatString, "Moving Weekly Total vs Moving Weekly Total -1 month", MWTvsMWTminus1label},
        {"MWT vs MWT-1(%)", MWTvsMWTminus1pct, pctFormatString, "Moving Weekly Total vs Moving Weekly Total -1 week (%)", MWTvsMWTminus1pctlabel}
    };


            int j = 0;


            //create calculation items for each calculation with formatstring and description
            foreach (var cg in Model.CalculationGroups)
            {
                if (cg.Name == calcGroupName)
                {
                    for (j = 0; j < calcItems.GetLength(0); j++)
                    {

                        string itemName = calcItems[j, 0];

                        string itemExpression = calcItemProtection.Replace("<CODE>", calcItems[j, 1]);
                        itemExpression = itemExpression.Replace("<LABELCODE>", calcItems[j, 4]);

                        string itemFormatExpression = calcItemFormatProtection.Replace("<CODE>", calcItems[j, 2]);
                        itemFormatExpression = itemFormatExpression.Replace("<LABELCODEFORMATSTRING>", "\"\"\"\" & " + calcItems[j, 4] + " & \"\"\"\"");

                        //if(calcItems[j,2] != defFormatString) {
                        //    itemFormatExpression = calcItemFormatProtection.Replace("<CODE>",calcItems[j,2]);
                        //};

                        string itemDescription = calcItems[j, 3];

                        if (!cg.CalculationItems.Contains(itemName))
                        {
                            var nCalcItem = cg.AddCalculationItem(itemName, itemExpression);
                            nCalcItem.FormatStringExpression = itemFormatExpression;
                            nCalcItem.FormatDax();
                            nCalcItem.Ordinal = j;
                            nCalcItem.Description = itemDescription;

                        };




                    };


                };
            };


        }

        void createMeasureFromCalcGroupWithFieldParameter()
        {
            ////uncoment the following three lines in TabularEditor
            //#r "Microsoft.VisualBasic"
            //using Microsoft.VisualBasic;
            //using System.Windows.Forms;

            /* '2023-01-26 / B.Agullo / creates a field parameter of measures filtered by calc group and values of a column with a name defined by a measure evaluated in the filtered value and calc item  */

            /* DYNAMIC HEADER FIELD PARAMETER SCRIPT */

            /* select measures and execute, you will need to run it twice */
            /* first time to create aux calc group, second time to actually create measuree*/
            /* remove aux calc group before going to production, do the right thing */
            

            string auxCgTag = "@AgulloBernat";
            string auxCgTagValue = "CG to extract format strings";

            string auxCalcGroupName = "DELETE AUX CALC GROUP";
            string auxCalcItemName = "Get Format String";

            string baseMeasureAnnotationName = "Base Measure";
            string calcItemAnnotationName = "Calc Item";
            string calcItemSortOrderName = "Sort Order";
            string calcItemSortOrderValue = String.Empty;

            string filterValueAnnotationName = String.Empty;
            string dynamicNameAnnotationName = "Dynamic Name";


            string scriptAnnotationName = "Script";
            string scriptAnnotationValue = "Create Measures with a Calculation Group "+ DateTime.Now.ToString("yyyyMMddHHmmss");

            bool generateFieldParameter;

            DialogResult dialogResult = MessageBox.Show("Generate Field Parameter?", "Field Parameter", MessageBoxButtons.YesNo);
            generateFieldParameter = (dialogResult == DialogResult.Yes);


            


            /*check if any measures are selected*/
            if (Selected.Measures.Count == 0)
            {
                Error("No measures selected");
                return;
            }

            /*find any regular CGs (excluding the one we might have created)*/
            var regularCGs = Model.Tables.Where(
                x => x.ObjectType == ObjectType.CalculationGroupTable
                & x.GetAnnotation(auxCgTag) != auxCgTagValue);

            if (regularCGs.Count() == 0)
            {
                Error("No Calculation Groups Found");
                return;
            };


            

            //the lambda expression will be avaluated for all calc groups to find a matching calc group
            //CalculationGroupTable auxCg = Fx.SelectCalculationGroup(model:Model,lambdaExpression:lambda,selectFirst:true, showErrorIfNoTablesFound:false);

            bool calcGroupWasCreated = false;
            
            //the calc group will only be created if not found, and when so the boolean will point to it
            CalculationGroupTable auxCg = Fx.AddCalculationGroupExt(model: Model, calcGroupWasCreated: out calcGroupWasCreated, 
                defaultName: auxCalcGroupName, customCalcGroupName: false, annotationName: auxCgTag, annotationValue: auxCgTagValue);

            if (calcGroupWasCreated)
            {
                CalculationItem cItem = Fx.AddCalculationItemExt(cg: auxCg, calcItemName: auxCalcItemName, valueExpression: "SELECTEDMEASUREFORMATSTRING()");
                auxCg.IsHidden = true; 
                
                Info("Save changes to the model, recalculate the model, and launch the script again.");
                return;
            }

            //to avoid showing the aux calc group in the list
            Func<Table, bool> lambda = (x) => x.GetAnnotation(auxCgTag) != auxCgTagValue;

            CalculationGroupTable regularCg = Fx.SelectCalculationGroup(model: Model, lambdaExpression: lambda);
            if (regularCg == null) return;


            Table filterTable = Fx.SelectTableExt(model: Model, excludeCalcGroups: true, label:"Select table of filter field",showErrorIfNoSelection:true);
            if(filterTable == null) return;
            Column filterColumn = SelectColumn(filterTable,label:"Select filter Field");
            if (filterColumn == null) return;

            filterValueAnnotationName = filterColumn.Name; 

            String filterQuery = String.Format("EVALUATE DISTINCT({0})", filterColumn.DaxObjectFullName);

            List<String> filterValues = new List<String>();

            using (var filterReader = Model.Database.ExecuteReader(filterQuery))
            {

                while (filterReader.Read())
                {

                    filterValues.Add(filterReader.GetValue(0).ToString());
                }
            }
            
            string name = String.Empty;
            if (generateFieldParameter)
            {
                name = Interaction.InputBox("Provide a name for the field parameter", "Field Parameter", regularCg.Name + " Measures", 740, 400);
                if (name == "") { Error("Execution Aborted"); return; };
            };

            Measure dynamicNameMeasure = SelectMeasure(label: "Select measure for dynamic name, cancel if none");


            /*iterates through each selected measure*/
            foreach (Measure m in Selected.Measures)
            {
                /*check that base measure has a proper format string*/
                if (m.FormatString == "")
                {
                    Error("Define FormatString for " + m.Name + " and try again");
                    return;
                };

                /*prepares a displayfolder to store all new measures*/
                string displayFolderName = m.Name + " Measures";

                /*iterates thorough all calculation items of the selected calc group*/
                foreach (CalculationItem calcItem in regularCg.CalculationItems)
                {

                    string measureNamePrefix = string.Concat(Enumerable.Repeat("\u8203", calcItem.Ordinal));

                    foreach (string filterValue in filterValues)
                    {
                        
                        
                        
                        /*measure name*/
                        string measureName = measureName = m.Name + " " + calcItem.Name + " " + filterValue;

                        string dynamicMeasureName = String.Empty;  

                        if (dynamicNameMeasure == null)
                        {
                            dynamicMeasureName = measureName;
                        }
                        else
                        {

                            string measureNameQuery = String.Empty;

                            if (filterColumn.DataType == DataType.String)
                            {

                                measureNameQuery =
                                    String.Format("EVALUATE {{CALCULATE({0},{1}=\"{2}\",{3}=\"{4}\") & \"\"}}", 
                                        dynamicNameMeasure.DaxObjectFullName, 
                                        filterColumn.DaxObjectFullName, 
                                        filterValue,
                                        regularCg.Columns[0].DaxObjectFullName,
                                        calcItem.Name);
                            }
                            else
                            {
                                measureNameQuery =
                                    String.Format("EVALUATE {{CALCULATE({0},{1}={2},{3}=\"{4}\") & \"\"}}",
                                        dynamicNameMeasure.DaxObjectFullName,
                                        filterColumn.DaxObjectFullName,
                                        filterValue,
                                        regularCg.Columns[0].DaxObjectFullName,
                                        calcItem.Name);
                            }




                            using (var reader = Model.Database.ExecuteReader(measureNameQuery))
                            {
                                while (reader.Read())
                                {
                                    dynamicMeasureName = reader.GetString(0).ToString();

                                }
                            }

                            dynamicMeasureName = measureNamePrefix + dynamicMeasureName;
                        }

                        

                        //only if the measure is not yet there (think of reruns)
                        if (!Model.AllMeasures.Any(x => x.Name == measureName))
                        {

                            /*prepares a query to calculate the resulting format when applying the calculation item on the measure*/
                            string query = string.Format(
                                "EVALUATE {{CALCULATE({0},{1},{2})}}",
                                m.DaxObjectFullName,
                                string.Format(
                                    "{0}=\"{1}\"",
                                    regularCg.Columns[0].DaxObjectFullName,
                                    calcItem.Name),
                                string.Format(
                                    "{0}=\"{1}\"",
                                    auxCg.Columns[0].DaxObjectFullName,
                                    auxCalcItemName)
                            );

                            /*executes the query*/
                            using (var reader = Model.Database.ExecuteReader(query))
                            {
                                // resultset should contain just one row, with the format string
                                while (reader.Read())
                                {


                                    /*retrive the formatstring from the query*/
                                    string formatString = reader.GetValue(0).ToString();

                                    Output(formatString);




                                    /*build the expression of the measure*/
                                    string measureExpression = String.Empty;

                                    if(filterColumn.DataType == DataType.String)
                                    {
                                        measureExpression = string.Format(
                                            "CALCULATE({0},{1}=\"{2}\",KEEPFILTERS({3}=\"{4}\"))",
                                            m.DaxObjectName,
                                            regularCg.Columns[0].DaxObjectFullName,
                                            calcItem.Name,
                                            filterColumn.DaxObjectFullName,
                                            filterValue
                                        );
                                    }
                                    else
                                    {
                                        measureExpression = string.Format(
                                            "CALCULATE({0},{1}=\"{2}\",KEEPFILTERS({3}={4}))",
                                            m.DaxObjectName,
                                            regularCg.Columns[0].DaxObjectFullName,
                                            calcItem.Name,
                                            filterColumn.DaxObjectFullName,
                                            filterValue
                                        );
                                    }

                                        
                                        
                                        



                                    /*actually build the measure*/
                                    Measure newMeasure =
                                        m.Table.AddMeasure(
                                            name: measureName,
                                            expression: measureExpression);


                                    /*the all important format string!*/
                                    newMeasure.FormatString = formatString;

                                    /*final polish*/
                                    newMeasure.DisplayFolder = displayFolderName;
                                    newMeasure.FormatDax();

                                    /*add annotations for the creation of the field parameter*/
                                    newMeasure.SetAnnotation(baseMeasureAnnotationName, m.Name);
                                    newMeasure.SetAnnotation(calcItemAnnotationName, calcItem.Name);
                                    newMeasure.SetAnnotation(scriptAnnotationName, scriptAnnotationValue);
                                    newMeasure.SetAnnotation(calcItemSortOrderName, calcItem.Ordinal.ToString("000"));
                                    newMeasure.SetAnnotation(filterValueAnnotationName, filterValue);
                                    newMeasure.SetAnnotation(dynamicNameAnnotationName, dynamicMeasureName);


                                }
                            }
                        }
                    }
                        
                    
                }
            }


            if (!generateFieldParameter)
            {
                //end of execution
                return;
            };


            // Before running the script, select the measures or columns that you
            // would like to use as field parameters (hold down CTRL to select multiple
            // objects). Also, you may change the name of the field parameter table
            // below. NOTE: If used against Power BI Desktop, you must enable unsupported
            // features under File > Preferences (TE2) or Tools > Preferences (TE3).


            if (Selected.Columns.Count == 0 && Selected.Measures.Count == 0) throw new Exception("No columns or measures selected!");

            // Construct the DAX for the calculated table based on the measures created previously by the script
            var objects = Model.AllMeasures
                .Where(x => x.GetAnnotation(scriptAnnotationName) == scriptAnnotationValue)
                .OrderBy(x => x.GetAnnotation(baseMeasureAnnotationName) + x.GetAnnotation(calcItemSortOrderName));

            var dax = "{\n    " + string.Join(",\n    ", objects.Select((c, i) => string.Format("(\"{6}\", NAMEOF('{1}'[{0}]), {2},\"{3}\",\"{4}\",\"{5}\")",
                c.Name, c.Table.Name, i,
                Model.Tables[c.Table.Name].Measures[c.Name].GetAnnotation(baseMeasureAnnotationName),
                Model.Tables[c.Table.Name].Measures[c.Name].GetAnnotation(calcItemAnnotationName),
                Model.Tables[c.Table.Name].Measures[c.Name].GetAnnotation(filterValueAnnotationName),
                Model.Tables[c.Table.Name].Measures[c.Name].GetAnnotation(dynamicNameAnnotationName)
                ))) + "\n}";

            // Add the calculated table to the model:
            var table = Model.AddCalculatedTable(name, dax);

            // In TE2 columns are not created automatically from a DAX expression, so 
            // we will have to add them manually:
            var te2 = table.Columns.Count == 0;
            var nameColumn = te2 ? table.AddCalculatedTableColumn(name, "[Value1]") : table.Columns["Value1"] as CalculatedTableColumn;
            var fieldColumn = te2 ? table.AddCalculatedTableColumn(name + " Fields", "[Value2]") : table.Columns["Value2"] as CalculatedTableColumn;
            var orderColumn = te2 ? table.AddCalculatedTableColumn(name + " Order", "[Value3]") : table.Columns["Value3"] as CalculatedTableColumn;

            if (!te2)
            {
                // Rename the columns that were added automatically in TE3:
                nameColumn.IsNameInferred = false;
                nameColumn.Name = name;
                fieldColumn.IsNameInferred = false;
                fieldColumn.Name = name + " Fields";
                orderColumn.IsNameInferred = false;
                orderColumn.Name = name + " Order";
            }
            // Set remaining properties for field parameters to work
            // See: https://twitter.com/markbdi/status/1526558841172893696
            nameColumn.SortByColumn = orderColumn;
            nameColumn.GroupByColumns.Add(fieldColumn);
            fieldColumn.SortByColumn = orderColumn;
            fieldColumn.SetExtendedProperty("ParameterMetadata", "{\"version\":3,\"kind\":2}", ExtendedPropertyType.Json);
            fieldColumn.IsHidden = true;
            orderColumn.IsHidden = true;
        }


        void copyMacroFromVSFile()
        {
            //#r "System.IO"
            //#r "Microsoft.CodeAnalysis"
            //using System.IO;
            //using System.Windows.Forms;
            //using Microsoft.CodeAnalysis;
            //using Microsoft.CodeAnalysis.CSharp;
            //using Microsoft.CodeAnalysis.CSharp.Syntax;

            // '2023-05-06 / B.Agullo / 
            // this macro copies the code of any of the methods defined in the TE_Scripts.cs File
            // if the macro is using the custom class it must include de following commented directive
            //           //using GeneralFunctions;
            // if this line is found the macro will copy the code also from the class defined in GeneralFunctions
            // and will combine the commented references of the class with those of the macro
            // once the macro finishes the code is in the clipboard so it can be pasted
            // in a new c# script tab in Tabular Editor, using CTRL+V 
            // see further detail at -- 

            //config
            String macroFilePath = @"<HERE FULL PATH TO TE_Scripts.cs FILE>";
            String customClassFilePath = @"<HERE FULL PATH TO GeneralFunctions.cs FILE>";
            String codeIndent = "            ";
            String customClassEndMark = @"//******************";

            //get file structure
            SyntaxTree tree = CSharpSyntaxTree.ParseText(File.ReadAllText(macroFilePath));

            //extract method names that are not public static (just macro names) 
            List<string> macroNames = tree.GetRoot().DescendantNodes().OfType<MethodDeclarationSyntax>()
                                            .Where(m => m.Modifiers.ToString() != "public static")
                                            .Select(m => m.Identifier.ToString()).ToList();

            // Code that defines a local function "SelectString", which pops up a listbox allowing the user to select a 
            // string from a number of options:
            Func<IList<string>, string, string> SelectString = (IList<string> options, string title) =>
            {
                var form = new Form();
                form.Text = title;
                var buttonPanel = new Panel();
                buttonPanel.Dock = DockStyle.Bottom;
                buttonPanel.Height = 30;
                var okButton = new Button() { DialogResult = DialogResult.OK, Text = "OK" };
                var cancelButton = new Button() { DialogResult = DialogResult.Cancel, Text = "Cancel", Left = 80 };
                var listbox = new ListBox();
                listbox.Dock = DockStyle.Fill;
                listbox.Items.AddRange(options.ToArray());
                listbox.SelectedItem = options[0];

                form.Controls.Add(listbox);
                form.Controls.Add(buttonPanel);
                buttonPanel.Controls.Add(okButton);
                buttonPanel.Controls.Add(cancelButton);

                var result = form.ShowDialog();
                if (result == DialogResult.Cancel) return null;
                return listbox.SelectedItem.ToString();
            };

            //check that macros were found
            if (macroNames.Count == 0)
            {
                Error("No macros found in " + macroFilePath);
                return;
            }

            //let the user select the name of the macro to copy
            String select = SelectString(macroNames, "Choose a macro");

            //check that indeed one macro was selected
            if (select == null)
            {
                Info("You cancelled!");
                return;
            }

            //get the method
            MethodDeclarationSyntax method = tree.GetRoot().DescendantNodes().OfType<MethodDeclarationSyntax>()
                .First(m => m.Identifier.ToString() == select);

            //fix the code
            String macroCode = method.Body.ToFullString().Replace("//using", "using").Replace("//#r", "#r");
            int firstCurlyBracket = macroCode.IndexOf("{");
            int lastCurlyBracket = macroCode.LastIndexOf("}");

            macroCode = macroCode.Substring(firstCurlyBracket + 1, lastCurlyBracket - firstCurlyBracket - 1);


            //check the custom className 
            SyntaxTree customClassTree = CSharpSyntaxTree.ParseText(File.ReadAllText(customClassFilePath));

            string customClassNamespaceName = customClassTree.GetRoot().DescendantNodes().OfType<NamespaceDeclarationSyntax>().First().Name.ToString();


            //check if macro is using custom class
            if (macroCode.Contains("using " + customClassNamespaceName))
            {

                ClassDeclarationSyntax customClass = customClassTree.GetRoot().DescendantNodes().OfType<ClassDeclarationSyntax>().First();

                String customClassCode = customClass.ToString();
                int endMarkIndex = customClassCode.IndexOf(customClassEndMark);

                //crop the last part and uncomment the closing bracket
                customClassCode = customClassCode.Substring(0, endMarkIndex - 1).Replace("//}", "}").Replace("//using", "using").Replace("//#r", "#r");


                int hashrFirstMacroCode = Math.Max(macroCode.IndexOf("#r"), 0);
                int hashrFirstCustomClass = customClassCode.IndexOf("#r");

                if (hashrFirstCustomClass != -1)
                {
                    int hashrLastCustomClass = customClassCode.LastIndexOf("#r");
                    int endOfHashrCustomClass = customClassCode.IndexOf(Environment.NewLine, hashrLastCustomClass);

                    string[] hashrLines = customClassCode.Substring(hashrFirstCustomClass, endOfHashrCustomClass - hashrFirstCustomClass).Split('\n');

                    foreach (String hashrLine in hashrLines)
                    {
                        //if #r directive not present
                        if (!macroCode.Contains(hashrLine))
                        {
                            //insert in the code right before the first one
                            macroCode = macroCode.Substring(0, Math.Max(hashrFirstMacroCode - 1, 0))
                                + hashrLine + Environment.NewLine
                                + macroCode.Substring(hashrFirstMacroCode);

                            //update the position of the first #r
                            hashrFirstMacroCode = Math.Max(customClassCode.IndexOf("#r"), 0);
                        }
                    }

                    //remove #r directives from custom class 
                    customClassCode = customClassCode.Replace(customClassCode.Substring(hashrLastCustomClass, endOfHashrCustomClass - hashrLastCustomClass), "");

                }

                int usingFirstMacroCode = Math.Max(macroCode.IndexOf("using"), 0);
                int usingFirstCustomClass = customClassCode.IndexOf("using");

                if (usingFirstCustomClass != -1)
                {
                    int usingLastCustomClass = customClassCode.LastIndexOf("using");
                    int endOfusingCustomClass = customClassCode.IndexOf(Environment.NewLine, usingLastCustomClass);

                    string[] usingLines = customClassCode.Substring(usingFirstCustomClass, endOfusingCustomClass - usingFirstCustomClass).Split('\n');

                    foreach (String usingLine in usingLines)
                    {


                        //if #r directive not present
                        if (!macroCode.Contains(usingLine))
                        {
                            //insert in the code right before the first one
                            macroCode = macroCode.Substring(0, Math.Max(usingFirstMacroCode - 1, 0))
                                + usingLine + Environment.NewLine
                                + macroCode.Substring(usingFirstMacroCode);

                            usingFirstMacroCode = Math.Max(macroCode.IndexOf("using"), 0);
                        }
                    }

                    //remove using directives from custom class 
                    customClassCode = customClassCode.Replace(customClassCode.Substring(usingFirstCustomClass, endOfusingCustomClass - usingFirstCustomClass), "");

                }

                //remove the using directive since it is an in-script custom class
                macroCode = macroCode.Replace("using " + customClassNamespaceName, "");

                //append custom class to macro 
                macroCode += customClassCode;
            }

            string macroCodeClean = "";
            string[] macroCodeLines = macroCode.Split('\n');
            foreach (string macroCodeLine in macroCodeLines)
            {
                if (macroCodeLine.StartsWith(codeIndent))
                {
                    macroCodeClean += macroCodeLine.Substring(codeIndent.Length);
                }
                else
                {
                    macroCodeClean += macroCodeLine;
                }
            }

            //copy the code to the clipboard
            Clipboard.SetText(macroCodeClean);


        }

        //these two are necessary to have the Model and Selected objects available in the script
        static readonly Model Model;
        static readonly TabularEditor.UI.UITreeSelection Selected;


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

        public static Table SelectTable(Table preselect = null, string label = "Select Table")
        {
            return ScriptHelper.SelectTable(preselect: preselect, label: label);
        }

        public static Measure SelectMeasure(Measure preselect = null, string label = "Select Measure")
        {
            return ScriptHelper.SelectMeasure(preselect:preselect,label:label);
        }

        public static Measure SelectMeasure(IEnumerable<Measure> measures, Measure preselect = null, string label = "Select Measure")
        {
            return ScriptHelper.SelectMeasure(measures:measures, preselect: preselect, label: label);
        }


        public static Column SelectColumn(Table table, Column preselect = null, string label = "Select Column")
        {
            return ScriptHelper.SelectColumn(table: table, preselect: preselect, label: label); 
        }

        public static Column SelectColumn(IEnumerable<Column> columns, Column preselect = null, string label = "Select Column")
        {
            return ScriptHelper.SelectColumn(columns: columns, preselect: preselect, label: label);
        }

        public static void Output(object value, int lineNumber = -1)
        {
            ScriptHelper.Output(value: value, lineNumber: lineNumber);
        }


    }
}
