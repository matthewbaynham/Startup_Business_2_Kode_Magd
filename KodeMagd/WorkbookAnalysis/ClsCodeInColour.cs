using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Reflection;

namespace KodeMagd.WorkbookAnalysis
{
    public class ClsCodeInColour
    {
        public const string sCssName_DeclareVariables = "CssDeclareVariables";
        public const string sCssName_AssignVariables = "CssAssignVariables";
        public const string sCssName_IfStatements = "CssIfStatements";
        public const string sCssName_Loops = "CssLoops";
        public const string sCssName_DeclareFunctions = "CssDeclareFunctions";
        public const string sCssName_Comments = "CssComments";
        public const string sCssName_Errors = "Errors";
        public const string sCssName_With = "With";


        public enum enumCodeColourType
        {
                eDeclaringVariables,
                eAssigningValues,
                eIfStatements,
                eLoops,
                eFunctions,
                eComments,
                eErrors,
                eWith,
                eUnknown
        }

        /*
        public struct strCssStyleItem
        {
            public string sName;
            public string sValue;
        }

        public struct strCssStyle 
        {
            public string sName;
            public List<strCssStyleItem> lstItems;
        }
        */
        public static enumCodeColourType convert(ClsCodeMapper.enumLineType eLineType)
        {
            try
            {
                enumCodeColourType eResult = enumCodeColourType.eUnknown;

                switch(eLineType)
                {
                
                    //Variable declaring stuff
                    case ClsCodeMapper.enumLineType.eLineType_Options:
                    case ClsCodeMapper.enumLineType.eLineType_DeInitialise:
                    case ClsCodeMapper.enumLineType.eLineType_Dim:
                    case ClsCodeMapper.enumLineType.eLineType_DllFunctionDeclare:
                    case ClsCodeMapper.enumLineType.eLineType_Initialise:
                        eResult = enumCodeColourType.eDeclaringVariables;
                        break;
                    //Assigning values
                    case ClsCodeMapper.enumLineType.eLineType_AssignValue:
                    case ClsCodeMapper.enumLineType.eLineType_Input:
                    case ClsCodeMapper.enumLineType.eLineType_Output:
                        eResult = enumCodeColourType.eAssigningValues;
                        break;
                    //if statements
                    case ClsCodeMapper.enumLineType.eLineType_If:
                    case ClsCodeMapper.enumLineType.eLineType_Else:
                    case ClsCodeMapper.enumLineType.eLineType_ElseIF:
                    case ClsCodeMapper.enumLineType.eLineType_ExitIf:
                    case ClsCodeMapper.enumLineType.eLineType_EndIf:
                        eResult = enumCodeColourType.eIfStatements;
                        break;
                        //loops
                    case ClsCodeMapper.enumLineType.eLineType_BeginLoop:
                    case ClsCodeMapper.enumLineType.eLineType_EndLoop:
                    case ClsCodeMapper.enumLineType.eLineType_ExitLoop:
                        eResult = enumCodeColourType.eLoops;
                        break;
                        //functions
                    case ClsCodeMapper.enumLineType.eLineType_Call:
                    case ClsCodeMapper.enumLineType.eLineType_FunctionName:
                    case ClsCodeMapper.enumLineType.eLineType_ExitFn:
                    case ClsCodeMapper.enumLineType.eLineType_EndFunction:
                        eResult = enumCodeColourType.eFunctions;
                        break;
                        //comments
                    case ClsCodeMapper.enumLineType.eLineType_Comment:
                        eResult = enumCodeColourType.eComments;
                        break;
                        //Errors
                    case ClsCodeMapper.enumLineType.eLineType_OnError:
                    case ClsCodeMapper.enumLineType.eLineType_Goto:
                    case ClsCodeMapper.enumLineType.eLineType_ContinuedFromAbove:
                    case ClsCodeMapper.enumLineType.eLineType_ErrorHandler:
                        eResult = enumCodeColourType.eErrors;
                        break;
                        //with
                    case ClsCodeMapper.enumLineType.eLineType_With:
                    case ClsCodeMapper.enumLineType.eLineType_EndWith:
                        eResult = enumCodeColourType.eWith;
                        break;
                    default:
                        //ClsCodeMapper.enumLineType.eLineType_Empty
                        //ClsCodeMapper.enumLineType.eLineType_Unknown
                        eResult = enumCodeColourType.eUnknown;
                        break;
                 }

                return eResult;
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

                return enumCodeColourType.eUnknown;
            }
        }


    }
}
