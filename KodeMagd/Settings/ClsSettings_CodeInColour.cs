using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Configuration;
using KodeMagd.WorkbookAnalysis;

namespace KodeMagd.Settings
{
    class ClsSettings_CodeInColour : ApplicationSettingsBase
    {
        /*
        ClsCodeInColour.enumCodeColourType.
                eDeclaringVariables,
                eAssigningValues,
                eIfStatements,
                eLoops,
                eFunctions,
                eComments,
                eErrors,
                eWith,
                eUnknown
        */

        [UserScopedSetting()]
        [DefaultSettingValue("#808000")]
        public string lineColour_AssignVariables
        {
            get { return ((string)this["lineColour_AssignVariables"]); }
            set { this["lineColour_AssignVariables"] = (string)value; }
        }

        [UserScopedSetting()]
        [DefaultSettingValue("#FF8040")]
        public string lineColour_DeclareVariables
        {
            get { return ((string)this["lineColour_DeclareVariables"]); }
            set { this["lineColour_DeclareVariables"] = (string)value; }
        }

        [UserScopedSetting()]
        [DefaultSettingValue("#008080")]
        public string lineColour_If
        {
            get { return ((string)this["lineColour_If"]); }
            set { this["lineColour_If"] = (string)value; }
        }

        [UserScopedSetting()]
        [DefaultSettingValue("#000000")]
        public string lineColour_Loops
        {
            get { return ((string)this["lineColour_Loops"]); }
            set { this["lineColour_Loops"] = (string)value; }
        }

        [UserScopedSetting()]
        [DefaultSettingValue("#FF0080")]
        public string lineColour_DeclareFunctions
        {
            get { return ((string)this["lineColour_DeclareFunctions"]); }
            set { this["lineColour_DeclareFunctions"] = (string)value; }
        }

        [UserScopedSetting()]
        [DefaultSettingValue("#00FF00")]
        public string lineColour_Comments
        {
            get { return ((string)this["lineColour_Comments"]); }
            set { this["lineColour_Comments"] = (string)value; }
        }

        [UserScopedSetting()]
        [DefaultSettingValue("#FF0000")]
        public string lineColour_Errors
        {
            get { return ((string)this["lineColour_Errors"]); }
            set { this["lineColour_Errors"] = (string)value; }
        }
        
        [UserScopedSetting()]
        [DefaultSettingValue("#ff8000")]
        public string lineColour_With
        {
            get { return ((string)this["lineColour_With"]); }
            set { this["lineColour_With"] = (string)value; }
        }

        [UserScopedSetting()]
        [DefaultSettingValue("#000000")]
        public string lineColour_Unknown
        {
            get { return ((string)this["lineColour_Unknown"]); }
            set { this["lineColour_Unknown"] = (string)value; }
        }
    }
}
