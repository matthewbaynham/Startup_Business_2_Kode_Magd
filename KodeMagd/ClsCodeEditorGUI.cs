using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using VBA = Microsoft.Vbe.Interop;
using System.Runtime.InteropServices;
using System.Reflection;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Core;
using KodeMagd.InsertCode;
using KodeMagd.Misc;
using KodeMagd.Rename;
using KodeMagd.Format;
using KodeMagd.License;
using KodeMagd.WorkbookAnalysis;
using KodeMagd.Dependencies;

namespace KodeMagd
{
    class ClsCodeEditorGUI
    {
        public const string csCommandBarName = "Kode Magd";
        private const string csCommandBarNameFaceId = "KodeMagd_FaceId";
        private const int ciIconsFacesToBeDisplayed = 100;

        Office.CommandBar cmd;
        Office.CommandBarButton btnAbout;

        Office.CommandBarPopup popReadability;
        Office.CommandBarButton btnAutoFormat;
        Office.CommandBarButton btnIndenting;
        Office.CommandBarButton btnSplitLines;
        Office.CommandBarButton btnSetLineLength;
        Office.CommandBarButton btnRemoveLineNo;
        Office.CommandBarButton btnDimSpacing;

        Office.CommandBarPopup popRename;
        Office.CommandBarButton btnRenameVariable;
        Office.CommandBarButton btnRenameFunction;
        Office.CommandBarButton btnRenameModule;

        Office.CommandBarPopup popSuggestions;
        Office.CommandBarButton btnGenerateReport;
        Office.CommandBarButton btnDetectUnused;

        Office.CommandBarPopup popAnalysis;
        Office.CommandBarPopup popDependencies;
        Office.CommandBarButton btnDependenciesVariable;
        Office.CommandBarButton btnDependenciesFunctionSub;
        Office.CommandBarButton btnDependenciesModClassForm;
        Office.CommandBarButton btnObjectModule;
        Office.CommandBarButton btnFlowDiagram;
        Office.CommandBarButton btnCodeInColour;

        Office.CommandBarPopup popInsertCode;
        Office.CommandBarPopup popInsertBasics;

        Office.CommandBarPopup popInsertDatabase;
        Office.CommandBarButton btnDB_GenerateConnectionString;
        Office.CommandBarButton btnDB_ConnectToDB;
        Office.CommandBarButton btnDB_ConnectToDBLoopThroughRST;
        Office.CommandBarButton btnDB_RunStoredProc;
        Office.CommandBarButton btnDB_UpdateInsertDelete;

        Office.CommandBarPopup popInsertForms;
        Office.CommandBarButton btnForm_OpenFormIncludingOpenargs;
        Office.CommandBarButton btnForm_PopulateListBoxComboBox;

        Office.CommandBarPopup popInsertFiles;
        Office.CommandBarButton btnFiles_LogClass;
        Office.CommandBarButton btnFiles_ReadWriteText;
        Office.CommandBarButton btnFiles_FileExists;
        Office.CommandBarButton btnFiles_ColumnHeaderTrackerClass;

        //Office.CommandBarPopup popInsertGraphs;
        //Office.CommandBarButton btnGraphics_Graphs;
        //Office.CommandBarButton btnGraphics_SparkLines;

        Office.CommandBarPopup popInsertMisc;
        Office.CommandBarButton btnMisc_Toolbars;
        Office.CommandBarButton btnMisc_ToolbarsIcons;
        Office.CommandBarButton btnMisc_Email;
        Office.CommandBarButton btnMisc_ErrorHandler;
        Office.CommandBarButton btnMisc_Class;
        Office.CommandBarButton btnMisc_CreatePivotTable;

        Office.CommandBarPopup popSettings;
        Office.CommandBarButton btnSettings;
        Office.CommandBarButton btnPayment;

        Office.CommandBarComboBox btnFaceId_Range_Cmb;
        //Office.CommandBarComboBox btnFaceId_To_Cmb;
        //Office.CommandBarButton btnFaceId_Refresh;
        List<Office.CommandBarButton> lstBtnFaceId = new List<CommandBarButton>();

        
        //VBA.WindowsClass.ToolWindow tw;
        VBA.Window win;
        
        public void createCommandBar()
        {
            try
            {
                //Icons (mostly the same as excel)
                //http://www.outlookexchange.com/articles/toddwalker/BuiltInOLKIcons.asp

                ClsIntellilock cLicense = new ClsIntellilock();


                /************************************
                 * Command Bar                      *
                 ************************************/

                bool bIsTrusted = true;
                try
                {
                    if (Globals.ThisAddIn.Application.VBE.CommandBars.Count == 0)
                    {
                        //do nothing
                    }
                    bIsTrusted = true;
                }
                catch
                {
                    MessageBox.Show(ClsMessages.csWarning_ExcelSettings, ClsDefaults.messageBoxTitle());
                    bIsTrusted = false;
                }

                if (bIsTrusted)
                {
                    if (isExistsCommandBar(csCommandBarName)) 
                    {
                        deleteCommandBar(csCommandBarName);
                    }

                    cmd = Globals.ThisAddIn.Application.VBE.CommandBars.Add(csCommandBarName, MsoBarPosition.msoBarTop, false, true);

                    cmd.Visible = true;
                    cmd.Enabled = true;


                    /************************************
                     * About                            *
                     ************************************/
                    btnAbout = (Office.CommandBarButton)cmd.Controls.Add(MsoControlType.msoControlButton);
                    btnAbout.Caption = ClsMisc.gcsAppName;
                    btnAbout.DescriptionText = "About " + ClsMisc.gcsAppName;
                    btnAbout.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    btnAbout.FaceId = 2950;
                    btnAbout.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnAbout_Click);

                    /************************************
                     * menu Readability                 *
                     ************************************/
                    popReadability = (Office.CommandBarPopup)cmd.Controls.Add(MsoControlType.msoControlPopup);
                    popReadability.Caption = "Readability";

                    //auto Format
                    btnAutoFormat = (Office.CommandBarButton)popReadability.Controls.Add(MsoControlType.msoControlButton);
                    btnAutoFormat.Caption = "Auto Format";
                    btnAutoFormat.DescriptionText = "Auto Format";
                    btnAutoFormat.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    btnAutoFormat.FaceId = 1032;
                    if (cLicense.locked)
                    { btnAutoFormat.Enabled = false; }
                    else
                    { btnAutoFormat.Enabled = true; }
                    btnAutoFormat.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnAutoFormat_Click);

                    //Indenting
                    btnIndenting = (Office.CommandBarButton)popReadability.Controls.Add(MsoControlType.msoControlButton);
                    btnIndenting.Caption = "Indenting";
                    btnIndenting.DescriptionText = "Indenting code";
                    btnIndenting.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    btnIndenting.FaceId = 503;
                    if (cLicense.locked)
                    { btnIndenting.Enabled = false; }
                    else
                    { btnIndenting.Enabled = true; }
                    btnIndenting.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnIndenting_Click);

                    //Split Lines
                    btnSplitLines = (Office.CommandBarButton)popReadability.Controls.Add(MsoControlType.msoControlButton);
                    btnSplitLines.Caption = "Split Lines";
                    btnSplitLines.DescriptionText = "Any Line with has : to put code on the same line will be split on to seperate lines";
                    btnSplitLines.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    btnSplitLines.FaceId = 2189;
                    if (cLicense.locked)
                    { btnSplitLines.Enabled = false; }
                    else
                    { btnSplitLines.Enabled = true; }
                    btnSplitLines.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnSplitLines_Click);

                    //Set Line
                    btnSetLineLength = (Office.CommandBarButton)popReadability.Controls.Add(MsoControlType.msoControlButton);
                    btnSetLineLength.Caption = "Set Line Length";
                    btnSetLineLength.DescriptionText = "Set Line Length";
                    btnSetLineLength.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    btnSetLineLength.FaceId = 2213;
                    if (cLicense.locked)
                    { btnSetLineLength.Enabled = false; }
                    else
                    { btnSetLineLength.Enabled = true; }
                    btnSetLineLength.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnSetLineLength_Click);

                    //Remove Line Numbers
                    btnRemoveLineNo = (Office.CommandBarButton)popReadability.Controls.Add(MsoControlType.msoControlButton);
                    btnRemoveLineNo.Caption = "Remove Line Numbers";
                    btnRemoveLineNo.DescriptionText = "Remove Line Numbers";
                    btnRemoveLineNo.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    btnRemoveLineNo.FaceId = 387;
                    if (cLicense.locked)
                    { btnRemoveLineNo.Enabled = false; }
                    else
                    { btnRemoveLineNo.Enabled = true; }
                    btnRemoveLineNo.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnRemoveLineNo_Click);

                    //Remove Line Numbers
                    btnDimSpacing = (Office.CommandBarButton)popReadability.Controls.Add(MsoControlType.msoControlButton);
                    btnDimSpacing.Caption = "Dim Spacing";
                    btnDimSpacing.DescriptionText = "Adjust the spacing between the Dim and the datatype";
                    btnDimSpacing.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    btnDimSpacing.FaceId = 313;
                    if (cLicense.locked)
                    { btnDimSpacing.Enabled = false; }
                    else
                    { btnDimSpacing.Enabled = true; }
                    btnDimSpacing.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnDimSpacing_Click);


                    /****************************************
                     * Workbook Dependencies                *
                     ****************************************/
                    popDependencies = (Office.CommandBarPopup)cmd.Controls.Add(MsoControlType.msoControlPopup);
                    popDependencies.Caption = "Dependencies";

                    //btnDependenciesVariable
                    btnDependenciesVariable = (Office.CommandBarButton)popDependencies.Controls.Add(MsoControlType.msoControlButton);
                    btnDependenciesVariable.Caption = "Variables";
                    btnDependenciesVariable.DescriptionText = "Generates a html report showing what the Dependencies are relating to a Variable";
                    btnDependenciesVariable.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    btnDependenciesVariable.FaceId = 206;
                    if (cLicense.locked)
                    { btnDependenciesVariable.Enabled = false; }
                    else
                    { btnDependenciesVariable.Enabled = true; }
                    btnDependenciesVariable.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnDependenciesVariable_Click);

                    //btnDependenciesFunctionSub
                    btnDependenciesFunctionSub = (Office.CommandBarButton)popDependencies.Controls.Add(MsoControlType.msoControlButton);
                    btnDependenciesFunctionSub.Caption = "Function/Sub/Property";
                    btnDependenciesFunctionSub.DescriptionText = "Generates a html report showing what the Dependencies are relating to a Function, Sub routine or property";
                    btnDependenciesFunctionSub.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    btnDependenciesFunctionSub.FaceId = 249;
                    if (cLicense.locked)
                    { btnDependenciesFunctionSub.Enabled = false; }
                    else
                    { btnDependenciesFunctionSub.Enabled = true; }
                    btnDependenciesFunctionSub.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnDependenciesFunctionSub_Click);

                    //btnDependenciesModClassForm
                    btnDependenciesModClassForm = (Office.CommandBarButton)popDependencies.Controls.Add(MsoControlType.msoControlButton);
                    btnDependenciesModClassForm.Caption = "Module/Class/Form";
                    btnDependenciesModClassForm.DescriptionText = "Generates a html report showing what the Dependencies are relating to a Module, Class or Form";
                    btnDependenciesModClassForm.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    btnDependenciesModClassForm.FaceId = 250;
                    if (cLicense.locked)
                    { btnDependenciesModClassForm.Enabled = false; }
                    else
                    { btnDependenciesModClassForm.Enabled = true; }
                    btnDependenciesModClassForm.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnDependenciesModClassForm_Click);

                    /****************************************
                     * Workbook Analysis                    *
                     ****************************************/
                    //Module Map
                    popAnalysis = (Office.CommandBarPopup)cmd.Controls.Add(MsoControlType.msoControlPopup);
                    popAnalysis.Caption = "Analysis";

                    //btnObjectModule
                    btnObjectModule = (Office.CommandBarButton)popAnalysis.Controls.Add(MsoControlType.msoControlButton);
                    btnObjectModule.Caption = "Object Model";
                    btnObjectModule.DescriptionText = "Generates a html report showing all the VBA Code modules, classes, subroutines, properties, etc";
                    btnObjectModule.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    btnObjectModule.FaceId = 303;
                    if (cLicense.locked)
                    { btnObjectModule.Enabled = false; }
                    else
                    { btnObjectModule.Enabled = true; }
                    btnObjectModule.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnObjectModule_Click);

                    ////btnFlowDiagram
                    //btnFlowDiagram = (Office.CommandBarButton)popAnalysis.Controls.Add(MsoControlType.msoControlButton);
                    //btnFlowDiagram.Caption = "Flow diagram";
                    //btnFlowDiagram.DescriptionText = "Generates a html report showing a flow diagram of function/sub/property";
                    //btnFlowDiagram.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    //btnFlowDiagram.FaceId = 190;
                    //if (cLicense.locked)
                    //{ btnFlowDiagram.Enabled = false; }
                    //else
                    //{ btnFlowDiagram.Enabled = true; }
                    //btnFlowDiagram.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnFlowDiagram_Click);

                    //btnCodeInColour
                    btnCodeInColour = (Office.CommandBarButton)popAnalysis.Controls.Add(MsoControlType.msoControlButton);
                    btnCodeInColour.Caption = "Code in Colour";
                    btnCodeInColour.DescriptionText = "Generates a html report showing VBA code in different colours to make it more readable.";
                    btnCodeInColour.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    btnCodeInColour.FaceId = 285;
                    if (cLicense.locked)
                    { btnCodeInColour.Enabled = false; }
                    else
                    { btnCodeInColour.Enabled = true; }
                    btnCodeInColour.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnCodeInColour_Click);

                    /************************************
                     * menu Rename                      *
                     ************************************/
                    popRename = (Office.CommandBarPopup)cmd.Controls.Add(MsoControlType.msoControlPopup);
                    popRename.Caption = "Rename";

                    //Rename Variable
                    btnRenameVariable = (Office.CommandBarButton)popRename.Controls.Add(MsoControlType.msoControlButton);
                    btnRenameVariable.Caption = "Variable";
                    btnRenameVariable.DescriptionText = "Rename Variable";
                    btnRenameVariable.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    btnRenameVariable.FaceId = 503;
                    if (cLicense.locked)
                    { btnRenameVariable.Enabled = false; }
                    else
                    { btnRenameVariable.Enabled = true; }
                    btnRenameVariable.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnRenameVariable_Click);

                    //Rename Function
                    btnRenameFunction = (Office.CommandBarButton)popRename.Controls.Add(MsoControlType.msoControlButton);
                    btnRenameFunction.Caption = "Function/Sub/Property";
                    btnRenameFunction.DescriptionText = "Rename Function / Sub / Property";
                    btnRenameFunction.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    btnRenameFunction.FaceId = 529;
                    if (cLicense.locked)
                    { btnRenameFunction.Enabled = false; }
                    else
                    { btnRenameFunction.Enabled = true; }
                    btnRenameFunction.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnRenameFunction_Click);

                    //Rename Module
                    btnRenameModule = (Office.CommandBarButton)popRename.Controls.Add(MsoControlType.msoControlButton);
                    btnRenameModule.Caption = "Module";
                    btnRenameModule.DescriptionText = "Rename Module";
                    btnRenameModule.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    btnRenameModule.FaceId = 684;
                    if (cLicense.locked)
                    { btnRenameModule.Enabled = false; }
                    else
                    { btnRenameModule.Enabled = true; }
                    btnRenameModule.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnRenameModule_Click);


                    /************************************
                     * menu Suggestions                 *
                     ************************************/
                    //popSuggestions = (Office.CommandBarPopup)cmd.Controls.Add(MsoControlType.msoControlPopup);
                    //popSuggestions.Caption = "Suggestions";

                    ////Generate Report
                    //btnGenerateReport = (Office.CommandBarButton)popSuggestions.Controls.Add(MsoControlType.msoControlButton);
                    //btnGenerateReport.Caption = "Generate Report";
                    //btnGenerateReport.DescriptionText = "Generate Report";
                    //btnGenerateReport.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    //btnGenerateReport.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnGenerateReport_Click);

                    ////btnDetectUnused
                    //btnDetectUnused = (Office.CommandBarButton)popSuggestions.Controls.Add(MsoControlType.msoControlButton);
                    //btnDetectUnused.Caption = "DetectUnused";
                    //btnDetectUnused.DescriptionText = "DetectUnused";
                    //btnDetectUnused.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    //btnDetectUnused.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnDetectUnused_Click);



                    /************************************
                     * menu Insert Code                 *
                     ************************************/
                    popInsertCode = (Office.CommandBarPopup)cmd.Controls.Add(MsoControlType.msoControlPopup);
                    popInsertCode.Caption = "Insert Code";

                    /* Database */
                    popInsertDatabase = (Office.CommandBarPopup)popInsertCode.Controls.Add(MsoControlType.msoControlPopup);
                    popInsertDatabase.Caption = "Database";

                    //btnDB_GenerateConnectionString
                    btnDB_GenerateConnectionString = (Office.CommandBarButton)popInsertDatabase.Controls.Add(MsoControlType.msoControlButton);
                    btnDB_GenerateConnectionString.Caption = "Generate Connection String";
                    btnDB_GenerateConnectionString.DescriptionText = "Generate Connection String";
                    btnDB_GenerateConnectionString.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    btnDB_GenerateConnectionString.FaceId = 346;
                    if (cLicense.locked)
                    { btnDB_GenerateConnectionString.Enabled = false; }
                    else
                    { btnDB_GenerateConnectionString.Enabled = true; }
                    btnDB_GenerateConnectionString.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnDB_GenerateConnectionString_Click);

                    /*
                    //btnDB_ConnectToDB
                    btnDB_ConnectToDB = (Office.CommandBarButton)popInsertDatabase.Controls.Add(MsoControlType.msoControlButton);
                    btnDB_ConnectToDB.Caption = "Connect To DB";
                    btnDB_ConnectToDB.DescriptionText = "Connect To DB";
                    btnDB_ConnectToDB.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    btnDB_ConnectToDB.FaceId = 2126;
                    if (cLicense.locked)
                    { btnDB_ConnectToDB.Enabled = false; }
                    else
                    { btnDB_ConnectToDB.Enabled = true; }
                    btnDB_ConnectToDB.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnDB_ConnectToDB_Click);
                    */

                    //btnDB_ConnectToDBLoopThroughRST
                    btnDB_ConnectToDBLoopThroughRST = (Office.CommandBarButton)popInsertDatabase.Controls.Add(MsoControlType.msoControlButton);
                    btnDB_ConnectToDBLoopThroughRST.Caption = "Create Recordset";
                    btnDB_ConnectToDBLoopThroughRST.DescriptionText = "Create Recordset";
                    btnDB_ConnectToDBLoopThroughRST.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    btnDB_ConnectToDBLoopThroughRST.FaceId = 2153;
                    if (cLicense.locked)
                    { btnDB_ConnectToDBLoopThroughRST.Enabled = false; }
                    else
                    { btnDB_ConnectToDBLoopThroughRST.Enabled = true; }
                    btnDB_ConnectToDBLoopThroughRST.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnDB_ConnectToDBLoopThroughRST_Click);

                    //btnDB_RunStoredProc
                    btnDB_RunStoredProc = (Office.CommandBarButton)popInsertDatabase.Controls.Add(MsoControlType.msoControlButton);
                    btnDB_RunStoredProc.Caption = "Run Stored Proc";
                    btnDB_RunStoredProc.DescriptionText = "Run Stored Proc";
                    btnDB_RunStoredProc.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    btnDB_RunStoredProc.FaceId = 2151;
                    if (cLicense.locked)
                    { btnDB_RunStoredProc.Enabled = false; }
                    else
                    { btnDB_RunStoredProc.Enabled = true; }
                    btnDB_RunStoredProc.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnDB_RunStoredProc_Click);

                    //btnDB_UpdateInsertDelete
                    btnDB_UpdateInsertDelete = (Office.CommandBarButton)popInsertDatabase.Controls.Add(MsoControlType.msoControlButton);
                    btnDB_UpdateInsertDelete.Caption = "Update/Insert/Delete";
                    btnDB_UpdateInsertDelete.DescriptionText = "Run the SQL statements for Update, Insert or Delete";
                    btnDB_UpdateInsertDelete.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    btnDB_UpdateInsertDelete.FaceId = 2010;
                    if (cLicense.locked)
                    { btnDB_UpdateInsertDelete.Enabled = false; }
                    else
                    { btnDB_UpdateInsertDelete.Enabled = true; }
                    btnDB_UpdateInsertDelete.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnDB_UpdateInsertDelete_Click);

                    /* Forms */
                    popInsertForms = (Office.CommandBarPopup)popInsertCode.Controls.Add(MsoControlType.msoControlPopup);
                    popInsertForms.Caption = "Forms";

                    //btnForm_OpenFormIncludingOpenargs
                    btnForm_OpenFormIncludingOpenargs = (Office.CommandBarButton)popInsertForms.Controls.Add(MsoControlType.msoControlButton);
                    btnForm_OpenFormIncludingOpenargs.Caption = "Open Form and Pass values to it";
                    btnForm_OpenFormIncludingOpenargs.DescriptionText = "Open Form and use Openargs to pass values to it";
                    btnForm_OpenFormIncludingOpenargs.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    btnForm_OpenFormIncludingOpenargs.FaceId = 2124;
                    if (cLicense.locked)
                    { btnForm_OpenFormIncludingOpenargs.Enabled = false; }
                    else
                    { btnForm_OpenFormIncludingOpenargs.Enabled = true; }
                    btnForm_OpenFormIncludingOpenargs.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnForm_OpenFormIncludingOpenargs_Click);

                    //btnForm_PopulateListBoxComboBox
                    btnForm_PopulateListBoxComboBox = (Office.CommandBarButton)popInsertForms.Controls.Add(MsoControlType.msoControlButton);
                    btnForm_PopulateListBoxComboBox.Caption = "Populate ListBox ComboBox";
                    btnForm_PopulateListBoxComboBox.DescriptionText = "Populate ListBox and ComboBox with values";
                    btnForm_PopulateListBoxComboBox.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    btnForm_PopulateListBoxComboBox.FaceId = 2529;
                    if (cLicense.locked)
                    { btnForm_PopulateListBoxComboBox.Enabled = false; }
                    else
                    { btnForm_PopulateListBoxComboBox.Enabled = true; }
                    btnForm_PopulateListBoxComboBox.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnForm_PopulateListBoxComboBox_Click);

                    /* Files */
                    popInsertFiles = (Office.CommandBarPopup)popInsertCode.Controls.Add(MsoControlType.msoControlPopup);
                    popInsertFiles.Caption = "Files";

                    //btnFiles_LogClass
                    btnFiles_LogClass = (Office.CommandBarButton)popInsertFiles.Controls.Add(MsoControlType.msoControlButton);
                    btnFiles_LogClass.Caption = "Create Log Class";
                    btnFiles_LogClass.DescriptionText = "Create Log Class";
                    btnFiles_LogClass.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    btnFiles_LogClass.FaceId = 627;
                    if (cLicense.locked)
                    { btnFiles_LogClass.Enabled = false; }
                    else
                    { btnFiles_LogClass.Enabled = true; }
                    btnFiles_LogClass.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnFiles_LogClass_Click);

                    //btnFiles_ReadWriteText
                    btnFiles_ReadWriteText = (Office.CommandBarButton)popInsertFiles.Controls.Add(MsoControlType.msoControlButton);
                    btnFiles_ReadWriteText.Caption = "Read Write Text File";
                    btnFiles_ReadWriteText.DescriptionText = "Read Write Text File";
                    btnFiles_ReadWriteText.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    btnFiles_ReadWriteText.FaceId = 5608;
                    if (cLicense.locked)
                    { btnFiles_ReadWriteText.Enabled = false; }
                    else
                    { btnFiles_ReadWriteText.Enabled = true; }
                    btnFiles_ReadWriteText.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnFiles_ReadWriteText_Click);

                    //btnFiles_FileExists
                    btnFiles_FileExists = (Office.CommandBarButton)popInsertFiles.Controls.Add(MsoControlType.msoControlButton);
                    btnFiles_FileExists.Caption = "File Exists";
                    btnFiles_FileExists.DescriptionText = "Check File Exists";
                    btnFiles_FileExists.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    btnFiles_FileExists.FaceId = 4087;
                    if (cLicense.locked)
                    { btnFiles_FileExists.Enabled = false; }
                    else
                    { btnFiles_FileExists.Enabled = true; }
                    btnFiles_FileExists.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnFiles_FileExists_Click);

                    //btnFiles_ColumnHeaderTrackerClass
                    btnFiles_ColumnHeaderTrackerClass = (Office.CommandBarButton)popInsertFiles.Controls.Add(MsoControlType.msoControlButton);
                    btnFiles_ColumnHeaderTrackerClass.Caption = "Column Header Tracker Class";
                    btnFiles_ColumnHeaderTrackerClass.DescriptionText = "A Class used to Track Column Headers in Input files";
                    btnFiles_ColumnHeaderTrackerClass.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    btnFiles_ColumnHeaderTrackerClass.FaceId = 2801;
                    if (cLicense.locked)
                    { btnFiles_ColumnHeaderTrackerClass.Enabled = false; }
                    else
                    { btnFiles_ColumnHeaderTrackerClass.Enabled = true; }
                    btnFiles_ColumnHeaderTrackerClass.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnFiles_ColumnHeaderTrackerClass_Click);

                    /*
            Office.CommandBarButton btnToolbars_Normal;
            Office.CommandBarButton btnToolbars_RightClick;
                     */

                    ///* popInsertGraphs */
                    //popInsertGraphs = (Office.CommandBarPopup)popInsertCode.Controls.Add(MsoControlType.msoControlPopup);
                    ////popInsertGraphs.= "Graphs";
                    //popInsertGraphs.Caption = "Graphs";

                    ////btnFiles_Graphs
                    //btnGraphics_Graphs = (Office.CommandBarButton)popInsertGraphs.Controls.Add(MsoControlType.msoControlButton);
                    //btnGraphics_Graphs.Caption = "Graphs";
                    //btnGraphics_Graphs.DescriptionText = "";
                    //btnGraphics_Graphs.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    //btnGraphics_Graphs.FaceId = 2519;
                    //btnGraphics_Graphs.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnFiles_Graphs_Click);

                    ////btnFiles_SparkLines
                    //btnGraphics_SparkLines = (Office.CommandBarButton)popInsertGraphs.Controls.Add(MsoControlType.msoControlButton);
                    //btnGraphics_SparkLines.Caption = "Sparklines";
                    //btnGraphics_SparkLines.DescriptionText = "";
                    //btnGraphics_SparkLines.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    //btnGraphics_SparkLines.FaceId = 2645;
                    //btnGraphics_SparkLines.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnFiles_SparkLines_Click);

                    /* popInsertMisc */
                    popInsertMisc = (Office.CommandBarPopup)popInsertCode.Controls.Add(MsoControlType.msoControlPopup);
                    popInsertMisc.Caption = "Misc";

                    //btnMisc_Toolbars
                    btnMisc_Toolbars = (Office.CommandBarButton)popInsertMisc.Controls.Add(MsoControlType.msoControlButton);
                    btnMisc_Toolbars.Caption = "Toolbar";
                    btnMisc_Toolbars.DescriptionText = "";
                    btnMisc_Toolbars.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    btnMisc_Toolbars.FaceId = 474;
                    if (cLicense.locked)
                    { btnMisc_Toolbars.Enabled = false; }
                    else
                    { btnMisc_Toolbars.Enabled = true; }
                    btnMisc_Toolbars.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnMisc_Toolbars_Click);

                    //btnMisc_ToolbarsIcons
                    btnMisc_ToolbarsIcons = (Office.CommandBarButton)popInsertMisc.Controls.Add(MsoControlType.msoControlButton);
                    btnMisc_ToolbarsIcons.Caption = "Toolbar Icons";
                    btnMisc_ToolbarsIcons.DescriptionText = "";
                    btnMisc_ToolbarsIcons.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    btnMisc_ToolbarsIcons.FaceId = 474;
                    if (cLicense.locked)
                    { btnMisc_ToolbarsIcons.Enabled = false; }
                    else
                    { btnMisc_ToolbarsIcons.Enabled = true; }
                    btnMisc_ToolbarsIcons.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnMisc_ToolbarsIcons_Click);

                    //btnMisc_Email
                    btnMisc_Email = (Office.CommandBarButton)popInsertMisc.Controls.Add(MsoControlType.msoControlButton);
                    btnMisc_Email.Caption = "Send Email";
                    btnMisc_Email.DescriptionText = "Send Email using Outlook";
                    btnMisc_Email.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    btnMisc_Email.FaceId = 5680;
                    if (cLicense.locked)
                    { btnMisc_Email.Enabled = false; }
                    else
                    { btnMisc_Email.Enabled = true; }
                    btnMisc_Email.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnMisc_Email_Click);

                    //btnMisc_ErrorHandler
                    btnMisc_ErrorHandler = (Office.CommandBarButton)popInsertMisc.Controls.Add(MsoControlType.msoControlButton);
                    btnMisc_ErrorHandler.Caption = "Error Handler";
                    btnMisc_ErrorHandler.DescriptionText = "Add Error Handler";
                    btnMisc_ErrorHandler.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    btnMisc_ErrorHandler.FaceId = 643;
                    if (cLicense.locked)
                    { btnMisc_ErrorHandler.Enabled = false; }
                    else
                    { btnMisc_ErrorHandler.Enabled = true; }
                    btnMisc_ErrorHandler.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnMisc_ErrorHandler_Click);

                    //btnMisc_Class
                    btnMisc_Class = (Office.CommandBarButton)popInsertMisc.Controls.Add(MsoControlType.msoControlButton);
                    btnMisc_Class.Caption = "New Class";
                    btnMisc_Class.DescriptionText = "Create new class";
                    btnMisc_Class.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    btnMisc_Class.FaceId = 609;
                    if (cLicense.locked)
                    { btnMisc_Class.Enabled = false; }
                    else
                    { btnMisc_Class.Enabled = true; }
                    btnMisc_Class.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnMisc_Class_Click);
                    /*
                    //btnMisc_CreatePivotTable
                    btnMisc_CreatePivotTable = (Office.CommandBarButton)popInsertMisc.Controls.Add(MsoControlType.msoControlButton);
                    btnMisc_CreatePivotTable.Caption = "Create Pivot Table Dynamically";
                    btnMisc_CreatePivotTable.DescriptionText = "Create Pivot Table Dynamically";
                    btnMisc_CreatePivotTable.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    btnMisc_CreatePivotTable.FaceId = 995;
                    if (cLicense.locked)
                    { btnMisc_CreatePivotTable.Enabled = false; }
                    else
                    { btnMisc_CreatePivotTable.Enabled = true; }
                    btnMisc_CreatePivotTable.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnMisc_CreatePivotTable_Click);
                    */

                    /************************************
                     * Settings                         *
                     ************************************/
                    popSettings = (Office.CommandBarPopup)cmd.Controls.Add(MsoControlType.msoControlPopup);
                    popSettings.Caption = "Settings...";

                    btnSettings = (Office.CommandBarButton)popSettings.Controls.Add(MsoControlType.msoControlButton);
                    btnSettings.Caption = "Settings";
                    btnSettings.DescriptionText = "General Settings";
                    btnSettings.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    btnSettings.FaceId = 3000;
                    btnSettings.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnSettings_Click);

                    btnPayment = (Office.CommandBarButton)popSettings.Controls.Add(MsoControlType.msoControlButton);
                    btnPayment.Caption = "License";
                    btnPayment.DescriptionText = "License details";
                    btnPayment.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    btnPayment.FaceId = 225;
                    btnPayment.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnPayment_Click);
                }

                cLicense = null;
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        public void deleteCommandBar(string sName)
        {
            try
            {
                if (this.isExistsCommandBar(sName))
                { 
                    try
                    { Globals.ThisAddIn.Application.VBE.CommandBars[sName].Delete();  }
                    catch
                    { MessageBox.Show(ClsMessages.csWarning_ExcelSettings, ClsDefaults.messageBoxTitle()); }
                }
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private bool isExistsCommandBar(string sName) 
        {
            try
            {
                bool bResult = false;

                foreach (Office.CommandBar cmdTemp in Globals.ThisAddIn.Application.VBE.CommandBars) 
                {
                    if (cmdTemp.Name.Trim().ToUpper() == sName.Trim().ToUpper()) 
                    { bResult = true; }
                }

                return bResult;
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);

                return false;
            }
        }

        private void btnAbout_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                //bool bIsOk = true;

                //if (!ClsMisc.checkCodeIsActive())
                //{
                //    MessageBox.Show("Please make sure that a code window is active.\n\nIf no code window is active there is nowhere to insert the newly generated code.");
                //    bIsOk = false;
                //}

                //if (bIsOk)
                //{
                    FrmAbout frm = new FrmAbout();

                    frm.ShowDialog();

                    frm = null;
                //}
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }


        private void btnAutoFormat_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                ClsCodeCleaner cCodeCleaner = new ClsCodeCleaner();

                cCodeCleaner.cleanModule(ClsCodeCleaner.enumCleaningType.eClean_All);

                cCodeCleaner = null;
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void btnIndenting_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                ClsCodeCleaner cCodeCleaner = new ClsCodeCleaner();
                
                cCodeCleaner.cleanModule(ClsCodeCleaner.enumCleaningType.eClean_Indenting);

                cCodeCleaner = null;
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }


        private void btnSplitLines_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                ClsCodeCleaner cCodeCleaner = new ClsCodeCleaner();

                cCodeCleaner.cleanModule(ClsCodeCleaner.enumCleaningType.eClean_SplitLines);

                cCodeCleaner = null;
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void btnSetLineLength_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                ClsCodeCleaner cCodeCleaner = new ClsCodeCleaner();

                cCodeCleaner.cleanModule(ClsCodeCleaner.enumCleaningType.eClean_SetLineLength);

                cCodeCleaner = null;
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void btnRemoveLineNo_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                FrmRemoveLineNo frm = new FrmRemoveLineNo();

                frm.ShowDialog();

                frm = null;
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void btnDimSpacing_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                ClsCodeCleaner cCodeCleaner = new ClsCodeCleaner();

                cCodeCleaner.cleanModule(ClsCodeCleaner.enumCleaningType.eClean_DimSpacing);

                cCodeCleaner = null;
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void btnDependenciesVariable_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                FrmDependenciesVariables frm = new FrmDependenciesVariables();

                frm.ShowDialog();

                frm = null;
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void btnDependenciesFunctionSub_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                FrmDependenciesFunction frm = new FrmDependenciesFunction();

                frm.ShowDialog();

                frm = null;
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void btnDependenciesModClassForm_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                FrmDependenciesModule frm = new FrmDependenciesModule();

                frm.ShowDialog();

                frm = null;
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }


        private void btnObjectModule_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                FrmObjectModel frm = new FrmObjectModel();

                frm.ShowDialog();

                frm = null;
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }


        private void btnFlowDiagram_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                FrmFlowDiagram frm = new FrmFlowDiagram();

                frm.ShowDialog();

                frm = null;
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void btnCodeInColour_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                FrmCodeInColour frm = new FrmCodeInColour();

                frm.ShowDialog();

                frm = null;
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }



        private void btnRenameVariable_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                FrmRenameVariable frm = new FrmRenameVariable();

                frm.ShowDialog();

                frm = null;

            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void btnRenameFunction_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                FrmRenameFunction frm = new FrmRenameFunction();

                frm.ShowDialog();

                frm = null;
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void btnRenameModule_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                FrmRenameModuleOrForm frm = new FrmRenameModuleOrForm();

                frm.ShowDialog();

                frm = null;
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void btnGenerateReport_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                MessageBox.Show("Generate Report", ClsDefaults.messageBoxTitle());
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void btnDetectUnused_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                MessageBox.Show("Detect Unused", ClsDefaults.messageBoxTitle());
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void btnDB_GenerateConnectionString_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                bool bIsOk = true;

                if (!ClsMisc.checkCodeIsActive())
                {
                    MessageBox.Show("Please make sure that a code window is active.\n\nIf no code window is active there is nowhere to insert the newly generated code.",
                                    ClsDefaults.messageBoxTitle(),
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Exclamation);
                    bIsOk = false;
                }

                if (bIsOk)
                {
                    FrmInsertCode_ConnectionString frm = new FrmInsertCode_ConnectionString();

                    frm.ShowDialog();

                    frm = null;
                }
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void btnDB_ConnectToDB_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                bool bIsOk = true;

                if (!ClsMisc.checkCodeIsActive())
                {
                    MessageBox.Show("Please make sure that a code window is active.\n\nIf no code window is active there is nowhere to insert the newly generated code.",
                                    ClsDefaults.messageBoxTitle(), 
                                    MessageBoxButtons.OK, 
                                    MessageBoxIcon.Exclamation);
                    bIsOk = false;
                }

                if (bIsOk)
                {
                    FrmOptions frm = new FrmOptions();

                    frm.ShowDialog();

                    frm = null;
                }
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void btnDB_ConnectToDBLoopThroughRST_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                bool bIsOk = true;

                if (!(ClsMisc.checkCodeIsActive() || ClsMisc.checkFormIsActive()))
                {
                    MessageBox.Show("Please make sure that a code window is active.\n\nIf no code window is active there is nowhere to insert the newly generated code.",
                                    ClsDefaults.messageBoxTitle(), 
                                    MessageBoxButtons.OK, 
                                    MessageBoxIcon.Exclamation);
                    bIsOk = false;
                }

                if (bIsOk)
                {
                    FrmInsertCode_Rst frm = new FrmInsertCode_Rst();

                    frm.ShowDialog();

                    frm = null;
                }
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void btnDB_RunStoredProc_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                bool bIsOk = true;

                if (!ClsMisc.checkCodeIsActive())
                {
                    MessageBox.Show("Please make sure that a code window is active.\n\nIf no code window is active there is nowhere to insert the newly generated code.",
                                    ClsDefaults.messageBoxTitle(), 
                                    MessageBoxButtons.OK, 
                                    MessageBoxIcon.Exclamation);
                    bIsOk = false;
                }

                if (bIsOk)
                {
                    FrmInsertCode_ExcuteStoredProcedure frm = new FrmInsertCode_ExcuteStoredProcedure();

                    frm.ShowDialog();
                    frm = null;
                }
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void btnDB_UpdateInsertDelete_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                bool bIsOk = true;

                if (!ClsMisc.checkCodeIsActive())
                {
                    MessageBox.Show("Please make sure that a code window is active.\n\nIf no code window is active there is nowhere to insert the newly generated code.",
                                    ClsDefaults.messageBoxTitle(),
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Exclamation);
                    bIsOk = false;
                }

                if (bIsOk)
                {
                    FrmInsertCode_SQL_UpdateInsertDelete frm = new FrmInsertCode_SQL_UpdateInsertDelete();

                    frm.ShowDialog();
                    frm = null;
                }
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void btnForm_OpenFormIncludingOpenargs_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                bool bIsOk = true;

                if (!ClsMisc.checkCodeIsActive())
                {
                    MessageBox.Show("Please make sure that a code window is active.\n\nIf no code window is active there is nowhere to insert the newly generated code.",
                                    ClsDefaults.messageBoxTitle(),
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Exclamation);
                    bIsOk = false;
                }

                if (bIsOk)
                {
                    FrmInsertCode_OpenForm frm = new FrmInsertCode_OpenForm();

                    frm.ShowDialog();

                    frm = null;
                }
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }


        private void btnForm_PopulateListBoxComboBox_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                bool bIsOk = true;

                if (!ClsMisc.checkCodeIsActive())
                {
                    MessageBox.Show("Please make sure that a code window is active.\n\nIf no code window is active there is nowhere to insert the newly generated code.",
                                    ClsDefaults.messageBoxTitle(),
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Exclamation);
                    bIsOk = false;
                }

                if (bIsOk)
                {
                    FrmInsertCode_Rst_PopulateListboxCombobox frm = new FrmInsertCode_Rst_PopulateListboxCombobox();

                    frm.ShowDialog();

                    frm = null;
                }
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void btnFiles_LogClass_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                bool bIsOk = true;

                if (!ClsMisc.checkCodeIsActive())
                {
                    MessageBox.Show("Please make sure that a code window is active.\n\nIf no code window is active there is nowhere to insert the newly generated code.",
                                    ClsDefaults.messageBoxTitle(),
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Exclamation);
                    bIsOk = false;
                }

                if (bIsOk)
                {
                    FrmInsertCode_TextFileLogClass frm = new FrmInsertCode_TextFileLogClass();

                    frm.ShowDialog();

                    frm = null;
                }
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void btnFiles_ReadWriteText_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                bool bIsOk = true;

                if (!ClsMisc.checkCodeIsActive())
                {
                    MessageBox.Show("Please make sure that a code window is active.\n\nIf no code window is active there is nowhere to insert the newly generated code.",
                                    ClsDefaults.messageBoxTitle(), 
                                    MessageBoxButtons.OK, 
                                    MessageBoxIcon.Exclamation);
                    bIsOk = false;
                }

                if (bIsOk)
                {
                    FrmInsertCode_TextFile frm = new FrmInsertCode_TextFile();

                    frm.ShowDialog();

                    frm = null;
                }
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void btnFiles_FileExists_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                bool bIsOk = true;

                if (!ClsMisc.checkCodeIsActive())
                {
                    MessageBox.Show("Please make sure that a code window is active.\n\nIf no code window is active there is nowhere to insert the newly generated code.",
                                    ClsDefaults.messageBoxTitle(),
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Exclamation);
                    bIsOk = false;
                }

                if (bIsOk)
                {
                    FrmInsertCode_FileExists frm = new FrmInsertCode_FileExists();

                    frm.ShowDialog();

                    frm = null;
                }
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void btnFiles_ColumnHeaderTrackerClass_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                bool bIsOk = true;

                if (!ClsMisc.checkCodeIsActive())
                {
                    MessageBox.Show("Please make sure that a code window is active.\n\nIf no code window is active there is nowhere to insert the newly generated code.",
                                    ClsDefaults.messageBoxTitle(),
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Exclamation);
                    bIsOk = false;
                }

                if (bIsOk)
                {
                    FrmInsertCode_ColumnPositionClass frm = new FrmInsertCode_ColumnPositionClass();

                    frm.ShowDialog();

                    frm = null;
                }
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void btnFiles_Graphs_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                MessageBox.Show("Code not yet written.");
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void btnFiles_SparkLines_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                bool bIsOk = true;

                if (!ClsMisc.checkCodeIsActive())
                {
                    MessageBox.Show("Please make sure that a code window is active.\n\nIf no code window is active there is nowhere to insert the newly generated code.",
                                    ClsDefaults.messageBoxTitle(), 
                                    MessageBoxButtons.OK, 
                                    MessageBoxIcon.Exclamation);
                    bIsOk = false;
                }

                if (bIsOk)
                {
                    FrmInsertCode_SparkLines frm = new FrmInsertCode_SparkLines();

                    frm.ShowDialog();

                    frm = null;
                }
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void btnMisc_Toolbars_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                bool bIsOk = true;

                if (!ClsMisc.checkCodeIsActive())
                {
                    MessageBox.Show("Please make sure that a code window is active.\n\nIf no code window is active there is nowhere to insert the newly generated code.",
                                    ClsDefaults.messageBoxTitle(),
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Exclamation);
                    bIsOk = false;
                }

                if (bIsOk)
                {
                    FrmInsertCode_CommandBarClass frm = new FrmInsertCode_CommandBarClass();

                    frm.ShowDialog();

                    frm = null;
                }
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void btnMisc_ToolbarsIcons_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                createCommandBarFaceId();
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void btnToolbars_RightClick_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                bool bIsOk = true;

                if (!ClsMisc.checkCodeIsActive())
                {
                    MessageBox.Show("Please make sure that a code window is active.\n\nIf no code window is active there is nowhere to insert the newly generated code.",
                                    ClsDefaults.messageBoxTitle(),
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Exclamation);
                    bIsOk = false;
                }

                if (bIsOk)
                {
                    FrmInsertCode_CommandBarClass frm = new FrmInsertCode_CommandBarClass();

                    frm.ShowDialog();

                    frm = null;
                }
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void btnMisc_Email_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                bool bIsOk = true;

                if (!ClsMisc.checkCodeIsActive())
                {
                    MessageBox.Show("Please make sure that a code window is active.\n\nIf no code window is active there is nowhere to insert the newly generated code.",
                                    ClsDefaults.messageBoxTitle(),
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Exclamation);
                    bIsOk = false;
                }

                if (bIsOk)
                {
                    FrmInsertCode_OutlookEMail frm = new FrmInsertCode_OutlookEMail();

                    frm.ShowDialog();

                    frm = null;
                }
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void btnMisc_ErrorHandler_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                bool bIsOk = true;

                if (!ClsMisc.checkCodeIsActive())
                {
                    MessageBox.Show("Please make sure that a code window is active.\n\nIf no code window is active there is nowhere to insert the newly generated code.",
                                    ClsDefaults.messageBoxTitle(),
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Exclamation);
                    bIsOk = false;
                }

                if (bIsOk)
                {
                    FrmInsertCode_ErrorHandler frm = new FrmInsertCode_ErrorHandler();

                    frm.ShowDialog();

                    frm = null;
                }
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void btnMisc_Class_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                bool bIsOk = true;

                if (!ClsMisc.checkCodeIsActive())
                {
                    MessageBox.Show("Please make sure that a code window is active.\n\nIf no code window is active there is nowhere to insert the newly generated code.",
                                    ClsDefaults.messageBoxTitle(),
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Exclamation);
                    bIsOk = false;
                }

                if (bIsOk)
                {
                    FrmInsertCode_Class frm = new FrmInsertCode_Class();

                    frm.ShowDialog();

                    frm = null;
                }
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void btnMisc_CreatePivotTable_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                bool bIsOk = true;

                if (!ClsMisc.checkCodeIsActive())
                {
                    MessageBox.Show("Please make sure that a code window is active.\n\nIf no code window is active there is nowhere to insert the newly generated code.",
                                    ClsDefaults.messageBoxTitle(),
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Exclamation);
                    bIsOk = false;
                }

                if (bIsOk)
                {
                    FrmInsertCode_PivotTable frm = new FrmInsertCode_PivotTable();

                    frm.ShowDialog();

                    frm = null;
                }
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void btnSettings_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                FrmOptions frm = new FrmOptions();

                frm.ShowDialog();

                frm = null;
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void btnPayment_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                FrmLicense frm = new FrmLicense();

                frm.ShowDialog();

                frm = null;
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        public void createCommandBarFaceId() 
        {
            try
            {
                if (isExistsCommandBar(csCommandBarNameFaceId))
                { deleteCommandBar(csCommandBarNameFaceId); }

                cmd = Globals.ThisAddIn.Application.VBE.CommandBars.Add(csCommandBarNameFaceId, MsoBarPosition.msoBarFloating, false, true);
                cmd.Visible = true;
                cmd.Enabled = true;
                cmd.Protection = MsoBarProtection.msoBarNoChangeDock;

                /*
                Office.CommandBarComboBox btnFaceId_From_Cmb;
                Office.CommandBarComboBox btnFaceId_To_Cmb;
                Office.CommandBarButton btnFaceId_Refresh;
                */
                btnFaceId_Range_Cmb = (Office.CommandBarComboBox)cmd.Controls.Add(MsoControlType.msoControlComboBox);
                btnFaceId_Range_Cmb.Caption = "Range";
                btnFaceId_Range_Cmb.DescriptionText = "Range";
                //btnFaceId_Range_Cmb.Style = MsoComboStyle.msoComboNormal;
                btnFaceId_Range_Cmb.Style = MsoComboStyle.msoComboLabel;

                btnFaceId_Range_Cmb.Text = "";

                for (int iCounter = 1; iCounter <= 90; iCounter++)
                {
                    int iFrom = ((iCounter - 1) * ciIconsFacesToBeDisplayed) + 1;
                    int iTo = (iCounter * ciIconsFacesToBeDisplayed);
                    string sItem = iFrom.ToString() + " to " + iTo.ToString();
                    btnFaceId_Range_Cmb.AddItem(sItem);

                    if (btnFaceId_Range_Cmb.Text == "")
                    { btnFaceId_Range_Cmb.Text = sItem; }
                }

                btnFaceId_Range_Cmb.Enabled = true;
                btnFaceId_Range_Cmb.Change += new Microsoft.Office.Core._CommandBarComboBoxEvents_ChangeEventHandler(btnFaceId_Refresh_Click);

                for (int iFaceId = 1; iFaceId < 100; iFaceId++)
                {
                    Office.CommandBarButton btnFaceId;

                    btnFaceId = (Office.CommandBarButton)cmd.Controls.Add(MsoControlType.msoControlButton);
                    btnFaceId.Caption = iFaceId.ToString();
                    btnFaceId.DescriptionText = iFaceId.ToString();
                    btnFaceId.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    btnFaceId.FaceId = iFaceId;

                    lstBtnFaceId.Add(btnFaceId);
                }

                cmd.Height = 400;
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void btnFaceId_Refresh_Click(Microsoft.Office.Core.CommandBarComboBox Ctrl)
        {
            try
            {
                bool bIsOk = true;
                int iRangeMin = 0;
                int iRangeMax = 0;
                string sText;
                
                if (btnFaceId_Range_Cmb.Text == null)
                { sText = ""; }
                else
                { sText = btnFaceId_Range_Cmb.Text.ToLower().Trim(); }

                if (!sText.Contains(" to "))
                { bIsOk = false; }

                if (bIsOk)
                {
                    int iPos = sText.IndexOf(" to ");

                    string sFrom = ClsMiscString.Left(sText, iPos);
                    string sTo = ClsMiscString.Right(sText, sText.Length - iPos - 4);

                    if (!int.TryParse(sFrom, out iRangeMin))
                    { bIsOk = false; }

                    if (!int.TryParse(sTo, out iRangeMax))
                    { bIsOk = false; }

                }

                if (bIsOk)
                {
                    int iHeight = cmd.Height;

                    foreach (CommandBarButton btn in lstBtnFaceId)
                    { cmd.Controls[btn.Index].Delete(); }
                    lstBtnFaceId.Clear();

                    for (int iFaceId = iRangeMin; iFaceId <= iRangeMax; iFaceId++)
                    {
                        Office.CommandBarButton btnFaceId;

                        btnFaceId = (Office.CommandBarButton)cmd.Controls.Add(MsoControlType.msoControlButton);
                        btnFaceId.Caption = iFaceId.ToString();
                        btnFaceId.DescriptionText = iFaceId.ToString();
                        btnFaceId.Style = MsoButtonStyle.msoButtonIconAndCaption;
                        btnFaceId.FaceId = iFaceId;

                        lstBtnFaceId.Add(btnFaceId);
                    }

                    cmd.Height = iHeight;
                }
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }
    }
}
