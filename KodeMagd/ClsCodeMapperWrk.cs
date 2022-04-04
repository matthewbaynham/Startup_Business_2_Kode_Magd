using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using VBA = Microsoft.Vbe.Interop;
using Microsoft.Vbe.Interop.Forms;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Text.RegularExpressions;
using KodeMagd.Misc;
using System.Diagnostics;
using KodeMagd.InsertCode;

namespace KodeMagd
{
    public class ClsCodeMapperWrk
    {
        List<ClsCodeMapper> lstCodeMapper = new List<ClsCodeMapper>();
        private Excel.Workbook objWrk;

        public struct strLinesInModule
        {
            public ClsCodeMapper.strModuleDetails objModuleDetails;
            public List<ClsCodeMapper.strLine> lstLines;
        }

        public Excel.Workbook Wrk 
        {
            get 
            {
                try{

                    return objWrk;
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
                    return null;
                }
            }
            set
            {
                try
                {
                    objWrk = value;

                    foreach (VBA.VBComponent vbComp in objWrk.VBProject.VBComponents)
                    {
                        ClsCodeMapper cCodeMapper = new ClsCodeMapper();
                        cCodeMapper.readCode(vbComp);
                        this.lstCodeMapper.Add(cCodeMapper);
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

        ~ClsCodeMapperWrk()
        {
            try
            {
                lstCodeMapper = null;
                objWrk = null;
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


        public List<String> getLstFunctionNames(bool bIncludeModuleNamePrefix)
        {
            try
            {
                List<string> lstResult = new List<string>();

                foreach (ClsCodeMapper cCodeMapper in lstCodeMapper)
                {
                    foreach (string sFunctionName in cCodeMapper.getLstFunctionNames())
                    {
                        if (bIncludeModuleNamePrefix)
                        { lstResult.Add(cCodeMapper.ModuleName + "." + sFunctionName); }
                        else
                        { lstResult.Add(sFunctionName); }
                    }
                }

                return lstResult;
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
                return null;
            }
        }

        public List<ClsCodeMapper.strFunctions> getLstFunctions()
        {
            try
            {
                List<ClsCodeMapper.strFunctions> lstResult = new List<ClsCodeMapper.strFunctions>();

                foreach (ClsCodeMapper cTemp in lstCodeMapper)
                {
                    string sModuleName = cTemp.ModuleDetails.sName;

                    foreach (ClsCodeMapper.strFunctions objTemp in getLstFunctions(sModuleName))
                    { lstResult.Add(objTemp); }
                }

                return lstResult;
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
                return null;
            }
        }


        public ClsCodeMapper.strFunctions getFunction(string sModuleName, string sFunctionName, ClsCodeMapper.enumFunctionType eFuntionType, ClsCodeMapper.enumFunctionPropertyType ePropertyType)
        {
            try
            {
                List<ClsCodeMapper.strFunctions> lstTemp = new List<ClsCodeMapper.strFunctions>();
                ClsCodeMapper.strFunctions objResult = new ClsCodeMapper.strFunctions();

                Predicate<ClsCodeMapper.strFunctions> predFunction;

                if (eFuntionType == ClsCodeMapper.enumFunctionType.eFnType_Property)
                { predFunction = x => x.sName.ToUpper().Trim() == sFunctionName.ToUpper().Trim() && x.ePropertyType == ePropertyType; }
                else
                { predFunction = x => x.sName.ToUpper().Trim() == sFunctionName.ToUpper().Trim(); }

                foreach (ClsCodeMapper cCodeMapper in lstCodeMapper)
                {
                    if (cCodeMapper.ModuleName.Trim().ToLower() == sModuleName.Trim().ToLower())
                    {
                        foreach (ClsCodeMapper.strFunctions objFunction in cCodeMapper.getLstFunctions().FindAll(predFunction))
                        { lstTemp.Add(objFunction); }
                    }
                }

                if (lstTemp.Count == 0)
                { }
                else if (lstTemp.Count > 1)
                { }
                else
                { objResult = lstTemp[0]; }

                return objResult;
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
                
                return new ClsCodeMapper.strFunctions();
            }
        }

        public List<ClsCodeMapper.strFunctions> getLstFunctions(string sModuleName)
        {
            try
            {
                List<ClsCodeMapper.strFunctions> lstResult = new List<ClsCodeMapper.strFunctions>();

                foreach (ClsCodeMapper cCodeMapper in lstCodeMapper)
                {
                    if (cCodeMapper.ModuleName.Trim().ToLower() == sModuleName.Trim().ToLower())
                    {
                        foreach (ClsCodeMapper.strFunctions objFunction in cCodeMapper.getLstFunctions())
                        { lstResult.Add(objFunction); }
                    }
                }

                return lstResult;
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
                return null;
            }
        }

        public List<String> getLstFunctionNames(string sModuleName, bool bIncludeModuleNamePrefix)
        {
            try
            {
                List<string> lstResult = new List<string>();

                foreach (ClsCodeMapper cCodeMapper in lstCodeMapper)
                {
                    if (cCodeMapper.ModuleName.Trim().ToLower() == sModuleName.Trim().ToLower())
                    {
                        foreach (string sFunctionName in cCodeMapper.getLstFunctionNames())
                        {
                            if (bIncludeModuleNamePrefix)
                            { lstResult.Add(cCodeMapper.ModuleName + "." + sFunctionName); }
                            else
                            { lstResult.Add(sFunctionName); }
                        }
                    }
                }

                return lstResult;
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
                return null;
            }
        }
        /*
        public string getActiveModuleName()
        {
            try
            {
                
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
                return null;
            }
        }
        */
        
        public bool moduleExists(string sName)
        {
            try
            {
                bool bResult = lstCodeMapper.Exists(x => x.ModuleDetails.sName.Trim().ToUpper() == sName.Trim().ToUpper());

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
        
        public List<ClsCodeMapper.strModuleDetails> getLstModuleDetails()
        {
            try
            {
                List<ClsCodeMapper.strModuleDetails> lstResult = new List<ClsCodeMapper.strModuleDetails>();

                foreach (ClsCodeMapper cCodeMapper in lstCodeMapper)
                { lstResult.Add(cCodeMapper.ModuleDetails); }

                return lstResult;
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
                return new List<ClsCodeMapper.strModuleDetails>();
            }
        }

        public List<ClsCodeMapper.strVariables> getLstVariableDetails()
        {
            try
            {
                List<ClsCodeMapper.strVariables> lstResult = new List<ClsCodeMapper.strVariables>();

                foreach (ClsCodeMapper cCodeMapper in lstCodeMapper)
                {
                    foreach (ClsCodeMapper.strVariables objVariable in cCodeMapper.lstVariables())
                    { lstResult.Add(objVariable); }
                }

                lstResult = lstResult.OrderBy(x => x.sName).ToList();

                return lstResult;
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
                return new List<ClsCodeMapper.strVariables>();
            }
        }

        public List<ClsCodeMapper.strVariables> getLstVariableDetails(string sModuleName)
        {
            try
            {
                List<ClsCodeMapper.strVariables> lstResult = new List<ClsCodeMapper.strVariables>();

                foreach (ClsCodeMapper cCodeMapper in lstCodeMapper.FindAll(x => x.ModuleDetails.sName.ToUpper().Trim() == sModuleName.ToUpper().Trim()))
                {
                    foreach (ClsCodeMapper.strVariables objVariable in cCodeMapper.lstVariables())
                    { lstResult.Add(objVariable); }
                }

                lstResult = lstResult.OrderBy(x => x.sName).ToList();

                return lstResult;
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
                return new List<ClsCodeMapper.strVariables>();
            }
        }

        public List<ClsCodeMapper.strVariables> getLstVariableDetails(string sModuleName, string sFunctionName, ClsCodeMapper.enumFunctionType eFunctionType, ClsCodeMapper.enumFunctionPropertyType ePropType)
        {
            try
            {
                List<ClsCodeMapper.strVariables> lstResult = new List<ClsCodeMapper.strVariables>();
                Predicate<ClsCodeMapper.strVariables> predVar;

                if (eFunctionType == ClsCodeMapper.enumFunctionType.eFnType_Property)
                { predVar = y => y.sFunctionName.ToUpper().Trim() == sFunctionName.ToUpper().Trim() && y.ePropType == ePropType; }
                else
                { predVar = y => y.sFunctionName.ToUpper().Trim() == sFunctionName.ToUpper().Trim(); }

                foreach (ClsCodeMapper cCodeMapper in lstCodeMapper.FindAll(x => x.ModuleDetails.sName.ToUpper().Trim() == sModuleName.ToUpper().Trim()))
                {
                    foreach (ClsCodeMapper.strVariables objVariable in cCodeMapper.lstVariables().FindAll(predVar))
                    { lstResult.Add(objVariable); }
                }

                lstResult = lstResult.OrderBy(x => x.sName).ToList();

                return lstResult;
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
                return new List<ClsCodeMapper.strVariables>();
            }
        }

        public ClsCodeMapper.strVariables getVariable(string sModuleName, string sFunctionName, ClsCodeMapper.enumFunctionPropertyType ePropType, string sVariableName)
        {
            try
            {
                /***********************
                 *   local Variables   *
                 ***********************/
                ClsCodeMapper.strVariables objResult = new ClsCodeMapper.strVariables();

                foreach (ClsCodeMapper cCodeMapper in lstCodeMapper.FindAll(y => y.ModuleDetails.sName.ToUpper().Trim() == sModuleName.ToUpper().Trim()))
                {
                    Predicate<ClsCodeMapper.strVariables> predVar;

                    if (ePropType == ClsCodeMapper.enumFunctionPropertyType.ePropType_NA)
                    { predVar = x => x.sFunctionName.ToUpper().Trim() == sFunctionName.ToUpper().Trim() && x.sName.ToUpper().Trim() == sVariableName.ToUpper().Trim(); }
                    else
                    { predVar = x => x.sFunctionName.ToUpper().Trim() == sFunctionName.ToUpper().Trim() && x.sName.ToUpper().Trim() == sVariableName.ToUpper().Trim() && x.ePropType == ePropType; }

                    objResult = cCodeMapper.lstVariables().Find(predVar);
                }

                return objResult;
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
                return new ClsCodeMapper.strVariables();
            }
        }

        public ClsCodeMapper.strVariables getVariable(string sModuleName, string sVariableName)
        {
            try
            {
                /************************
                 *   Global Variables   *
                 ************************/ 

                ClsCodeMapper.strVariables objResult = new ClsCodeMapper.strVariables();

                foreach (ClsCodeMapper cCodeMapper in lstCodeMapper.FindAll(y => y.ModuleDetails.sName.ToUpper().Trim() == sModuleName.ToUpper().Trim()))
                {
                    objResult = cCodeMapper.lstVariables().Find(x => x.eScope == ClsCodeMapper.enumScopeVar.eScope_Global || x.eScope == ClsCodeMapper.enumScopeVar.eScope_Module);
                }

                return objResult;
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
                return new ClsCodeMapper.strVariables();
            }
        }

        public bool functionNameExists(string sModuleName, string sName)
        {
            try
            {
                //List<string> lstFunctions = getLstFunctionNames(false);

                bool bResult = getLstFunctionNames(sModuleName, false).Contains(sName, StringComparer.OrdinalIgnoreCase);

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

        public bool functionNameExists(string sName)
        {
            try
            {
                //List<string> lstFunctions = getLstFunctionNames(false);

                bool bResult = getLstFunctionNames(false).Contains(sName, StringComparer.OrdinalIgnoreCase);

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

        public void renameFunctionInList(string sNewName, string sOldName, string sModuleName)
        {
            try
            {
                int iModIndex = lstCodeMapper.FindIndex(x => x.ModuleDetails.sName.Trim().ToLower() == sModuleName.Trim().ToLower());

                lstCodeMapper[iModIndex].renameFunctionInList(sNewName, sOldName);
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

        public void renameModuleInList(string sOldName, string sNewName)
        {
            try
            {
                int iModIndex = lstCodeMapper.FindIndex(x => x.ModuleDetails.sName.Trim().ToLower() == sOldName.Trim().ToLower());

                lstCodeMapper[iModIndex].renameModuleInList(sNewName);
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

        public void renameVariableInFunctionList(string sVariableNameOld, string sVariableNameNew, string sFunctionName, string sModuleName)
        {
            try
            {
                int iModIndex = lstCodeMapper.FindIndex(x => x.ModuleName.Trim().ToLower() == sModuleName.Trim().ToLower());

                if (iModIndex != -1)
                { lstCodeMapper[iModIndex].renameVariableInFunctionList(sVariableNameOld, sVariableNameNew, sFunctionName); }
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

        public bool variableNameExistInFunction(string sVariableName, string sFunctionName, string sModuleName)
        {
            try
            {
                if (sModuleName == null)
                { sModuleName = ""; }

                int iModIndex = lstCodeMapper.FindIndex(x => x.ModuleName.Trim().ToLower() == sModuleName.Trim().ToLower());
                bool bIsFound = false;

                if (iModIndex == -1)
                {
                    // not found
                }
                else
                { bIsFound = lstCodeMapper[iModIndex].variableNameExistsInFunction(sVariableName, sFunctionName); }

                return bIsFound;
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

        public void renameVariable(string sNewName, string sOldName)
        {
            try
            {
                foreach (ClsCodeMapper cTemp in lstCodeMapper)
                {
                    cTemp.RenameVariable(sNewName, sOldName);

                    cTemp.ImplementChanges();
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

        public List<ClsCodeMapperWrk.strLinesInModule> findModuleReferences(string sModuleName)
        {
            try
            {
                List<ClsCodeMapperWrk.strLinesInModule> lstResults = new List<ClsCodeMapperWrk.strLinesInModule>();

                foreach (ClsCodeMapper cCodeMapper in lstCodeMapper)
                {
                    ClsCodeMapperWrk.strLinesInModule objModuleDetails = new strLinesInModule();

                    objModuleDetails.objModuleDetails = cCodeMapper.ModuleDetails;
                    objModuleDetails.lstLines = new List<ClsCodeMapper.strLine>();

                    foreach (ClsCodeMapper.strLine objLine in cCodeMapper.findModuleReferences(sModuleName))
                    { objModuleDetails.lstLines.Add(objLine); }

                    lstResults.Add(objModuleDetails);
                }

                return lstResults;
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

                return new List<ClsCodeMapperWrk.strLinesInModule>();
            }
        }

        public void renameModule(string sNewName, string sOldName)
        {
            try
            {
                bool bIsOk = true;
                string sMessage = "";

                if (lstCodeMapper.Exists(x => x.ModuleDetails.sName.Trim().ToLower() == sNewName.Trim().ToLower()))
                {
                    bIsOk = false;
                    sMessage = "Module \"" + sNewName + "\" already exists.";
                }

                if (!lstCodeMapper.Exists(x => x.ModuleDetails.sName.Trim().ToLower() == sOldName.Trim().ToLower()))
                {
                    bIsOk = false;
                    sMessage = "The module your trying to change \"" + sOldName + "\" does not exist.";
                }

                if (bIsOk)
                { 
                    lstCodeMapper.Find(x => x.ModuleDetails.sName.Trim().ToLower() == sOldName.Trim().ToLower()).RenameModule(sNewName);

                    foreach (ClsCodeMapper cTemp in lstCodeMapper)
                    { 
                        cTemp.RenameVariable(sNewName, sOldName);
                        cTemp.ImplementChanges();
                    }
                }
                else
                { MessageBox.Show(sMessage, ClsDefaults.messageBoxTitle(), MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }

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

        public void renameFunction(string sNewName, string sOldName, string sModuleName)
        {
            try
            {


                //if public then change everywhere
                if (sModuleName == null)
                { sModuleName = ""; }

                bool bIsOk = true;
                string sMessage = "";
                ClsCodeMapper.enumScopeFn eScope = ClsCodeMapper.enumScopeFn.eScopeFn_Private;

                int iFnIndex = lstCodeMapper.FindIndex(x => x.ModuleDetails.sName.Trim().ToLower() == sModuleName.Trim().ToLower());

                if (iFnIndex == -1)
                {
                    bIsOk = false;
                    sMessage = "Can't find Function, Sub or Property";
                }
                else
                {
                    ClsCodeMapper.strFunctions objFn = lstCodeMapper[iFnIndex].getFunction(sOldName);

                    eScope = objFn.eScope;

                    ClsCodeMapper cCodeMapper = lstCodeMapper[iFnIndex];

                    cCodeMapper.RenameFunction(sNewName, sOldName);

                    if (eScope == ClsCodeMapper.enumScopeFn.eScopeFn_Private)
                    { 
                        cCodeMapper.RenameFunction(sNewName, sOldName);
                        cCodeMapper.ImplementChanges();
                    }
                    else
                    {
                        foreach (ClsCodeMapper cTemp in lstCodeMapper)
                        {
                            cTemp.RenameVariable(sNewName, sOldName);
                            cTemp.ImplementChanges();
                        }
                    }
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

        public void renameVariable(string sNewName, string sOldName, string sModuleName)
        {
            try
            {
                foreach (ClsCodeMapper cTemp in lstCodeMapper)
                {
                    if (cTemp.ModuleDetails.sName.Trim().ToLower() == sModuleName.Trim().ToLower())
                    {
                        cTemp.RenameVariable(sNewName, sOldName);
                        cTemp.ImplementChanges();
                    }
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

        public void renameVariable(string sNewName, string sOldName, string sModuleName, string sFunctionName)
        {
            try
            {
                if (sModuleName == null)
                { sModuleName = ""; }

                if (sFunctionName == null)
                { sFunctionName = ""; }

                int iModIndex = lstCodeMapper.FindIndex(x => x.ModuleDetails.sName.Trim().ToLower() == sModuleName.Trim().ToLower());

                if (iModIndex != -1)
                {
                    lstCodeMapper[iModIndex].RenameVariable(sNewName, sOldName, sFunctionName);
                    lstCodeMapper[iModIndex].ImplementChanges();
                }

                //foreach (ClsCodeMapper cCodeMapper in lstCodeMapper)
                //{
                //    if (cCodeMapper.ModuleDetails.sName.Trim().ToLower() == sModuleName.Trim().ToLower())
                //    {
                //        cCodeMapper.RenameVariable(sNewName, sOldName, sFunctionName);
                //        cCodeMapper.ImplementChanges();
                //    }
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

        public bool variableNameExistsGlobally(string sVariableName, bool bWithGlobalScope) 
        {
            try
            {
                bool bIsFound = false;

                foreach (ClsCodeMapper cCodeMapper in lstCodeMapper)
                {
                    if (!bIsFound & cCodeMapper.variableNameExistsInModule(sVariableName, bWithGlobalScope))
                    { bIsFound = true; }
                }

                return bIsFound;
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
        /*
        public bool moduleExists(string sModuleName)
        {
            try
            {
                bool bIsFound = false;

                foreach (ClsCodeMapper cCodeMapper in lstCodeMapper)
                {
                    if (cCodeMapper.ModuleDetails.sName.Trim().ToLower() == sModuleName.Trim().ToLower())
                    { bIsFound = true; }
                }

                return bIsFound;
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
        */
        public bool variableNameExistsInModule(string sVariableName, string sModuleName)
        {
            try
            {
                bool bIsFound = false;

                foreach (ClsCodeMapper cCodeMapper in lstCodeMapper)
                {
                    if (cCodeMapper.ModuleDetails.sName.Trim().ToLower() == sModuleName.Trim().ToLower())
                    {
                        if (cCodeMapper.variableNameExistsInModule(sVariableName, false))
                        { bIsFound = true; }
                    }
                }

                return bIsFound;
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

        public bool variableNameExists(string sVariableName, ClsCodeMapper.enumScopeVar eScope, ClsCodeMapper.strFunctions objFunction) 
        {
            try
            {
                bool bIsFound = false;
                
                foreach (ClsCodeMapper cCodeMapper in this.lstCodeMapper)
                {
                    ClsCodeMapper.strModuleDetails objModuleDetails = cCodeMapper.ModuleDetails;

                    if (objModuleDetails.sName.Trim().ToLower() == objFunction.sModuleName.Trim().ToLower())
                    {
                        bIsFound = true;
                        //cCodeMapper.
                    }
                }

                /*
                 * if eScope = enumScopeVar.eScope_Function then 
                 * check the function for local variables & function names
                 * check the module for global variables & function names
                  
                 * if the eScope = enumScopeVar.eScope_Module then
                 * and if module is this module check everything
                 * and if module is NOT this module check for global variables & function names
                  
                 * if eScope = enumScopeVar.eScope_Global then 
                 * check everything
                 
                 */
                /*
                eScope = enumScopeVar.eScope_Global

                bool bIsFound = false;

                foreach (strVariables objVar in lstVariablesMod)
                { 
                    if (objVar.sName.ToLower().Trim() == sNewName.ToLower().Trim())
                    {

                    }
                
                }

                int x = lstVariablesMod.FindIndex(x => x.sName.ToLower());

                sNewName.Trim().ToLower()




                foreach (strFunctions objFn in lstFunctions)
                { }
                */
                return false;
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
/*
        public ClsCodeMapper getModule(string sName)
        {
            try
            {
                lstCodeMapper.FindIndex(x => x.ModuleName.Trim().ToLower() == sName.Trim().ToLower());


                return false;
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
                return new ClsCodeMapper();
            }
        }
        */
        public bool existsGoToWithLineNo(string sFunctionName, string sModuleName)
        {
            bool bResult = false;

            try
            {
                if (lstCodeMapper.Exists(x => x.ModuleName.ToLower().Trim() == sModuleName.ToLower().Trim()))
                { bResult = lstCodeMapper.Find(x => x.ModuleName.ToLower().Trim() == sModuleName.ToLower().Trim()).existsGoToWithLineNo(sFunctionName); }

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
                return bResult;
            }
        }

        public void removeLineNo()
        {
            try
            {
                foreach (ClsCodeMapper cCodeMapper in lstCodeMapper)
                { cCodeMapper.removeLineNo(); }
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

        public void removeLineNo(string sFunctionName, string sModuleName)
        {
            try
            {
                bool bIsOK;

                if (sModuleName.Trim().ToLower() == ClsDefaults.textAll.Trim().ToLower())
                {
                    foreach (ClsCodeMapper objCodeMapper in lstCodeMapper)
                    { removeLineNo(sFunctionName, objCodeMapper.ModuleDetails.sName); }
                }
                else
                {
                    if (sModuleName.Trim().ToLower() == ClsDefaults.textAll.Trim().ToLower())
                    { bIsOK = true; }
                    else
                    {
                        if (lstCodeMapper.Exists(x => x.ModuleDetails.sName.Trim().ToLower() == sModuleName.Trim().ToLower()))
                        { bIsOK = true; }
                        else
                        { bIsOK = false; }
                    }

                    if (bIsOK)
                    {
                        int iIndex = lstCodeMapper.FindIndex(x => x.ModuleDetails.sName.Trim().ToLower() == sModuleName.Trim().ToLower());

                        lstCodeMapper[iIndex].removeLineNo(sFunctionName);
                        lstCodeMapper[iIndex].ImplementChanges();
                    }
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

        public void removeLineNo(string sModuleName)
        {
            try
            {
                bool bIsOK;

                if (sModuleName.Trim().ToLower() == ClsDefaults.textAll.Trim().ToLower())
                { bIsOK = true; }
                else
                {
                    if (lstCodeMapper.Exists(x => x.ModuleDetails.sName.Trim().ToLower() == sModuleName.Trim().ToLower()))
                    { bIsOK = true; }
                    else
                    { bIsOK = false; }
                }

                if (bIsOK)
                { 
                    int iIndex = lstCodeMapper.FindIndex(x => x.ModuleDetails.sName.Trim().ToLower() == sModuleName.Trim().ToLower());

                    lstCodeMapper[iIndex].removeLineNo();
                    lstCodeMapper[iIndex].ImplementChanges();
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

        //public void addErrorHandlerToFunction(string sModuleName, string sFunctionName, FrmInsertCode_ErrorHandler.enumReplaceErrorHandlerActions eActions) 
        //{
        //    try
        //    {
        //        if (lstCodeMapper.Exists(x => x.ModuleDetails.sName.Trim().ToLower() == sModuleName.Trim().ToLower()))
        //        { lstCodeMapper.Find(x => x.ModuleDetails.sName.Trim().ToLower() == sModuleName.Trim().ToLower()).addErrorHandlerToFunction(sFunctionName, eActions); }
        //    }
        //    catch (Exception ex)
        //    {
        //        MethodBase mbTemp = MethodBase.GetCurrentMethod();

        //        string sMessage = "";

        //        sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
        //        sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
        //        sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
        //        sMessage += ex.Message;

        //        MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
        //    }
        //}

        public List<ClsCodeMapper.strLine> getLines(string sModuleName)
        {
            try
            {
                if (lstCodeMapper.Exists(x => x.ModuleDetails.sName.Trim().ToLower() == sModuleName.Trim().ToLower()))
                {
                    return lstCodeMapper.Find(x => x.ModuleDetails.sName.Trim().ToLower() == sModuleName.Trim().ToLower()).lines;
                }
                else
                {
                    return new List<ClsCodeMapper.strLine>();
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

                return new List<ClsCodeMapper.strLine>();
            }
        }

        public List<ClsCodeMapper.strLine> getLines(string sModuleName, List<string> lstFunctionNames)
        {
            try
            {
                if (lstCodeMapper.Exists(x => x.ModuleDetails.sName.Trim().ToLower() == sModuleName.Trim().ToLower()))
                {
                    return lstCodeMapper.Find(x => x.ModuleDetails.sName.Trim().ToLower() == sModuleName.Trim().ToLower()).lines.FindAll(y => lstFunctionNames.Exists(z => z.ToLower().Trim() == y.sFunctionName.ToLower().Trim()));
                }
                else
                {
                    return new List<ClsCodeMapper.strLine>();
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

                return new List<ClsCodeMapper.strLine>();
            }
        }

        public ClsCodeMapper getCodeMapper(string sModuleName)
        {
            try
            {
                if (lstCodeMapper.Exists(x => x.ModuleDetails.sName.Trim().ToLower() == sModuleName.Trim().ToLower()))
                { return lstCodeMapper.Find(x => x.ModuleDetails.sName.Trim().ToLower() == sModuleName.Trim().ToLower()); }
                else
                { return new ClsCodeMapper(); }
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

                return new ClsCodeMapper();
            }
        }
    }
}
