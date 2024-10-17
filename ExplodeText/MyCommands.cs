using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using CadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(ExplodeText.MyCommands))]

namespace ExplodeText
{
    public class MyCommands 
    {
        [CommandMethod("BlowUpText")]
        public static void RunMyCommand()
        {
            var dwg = CadApp.DocumentManager.MdiActiveDocument;
            var editor = dwg.Editor;

            var txtId = SelectText(editor);
            if (!txtId.IsNull)
            {
                var exploder = new TextExploder();
                exploder.ExplodeText(dwg, txtId, "", true);
            }
            else
            {
                editor.WriteMessage("\n*Cancel*");
            }
            editor.PostCommandPrompt();
        }

        private static ObjectId SelectText(Editor ed)
        {
            var opt = new PromptEntityOptions(
                "\nSelect a TEXT/MTEXT entity:");
            opt.SetRejectMessage("\nInvalid: must be a TEXT entity.");
            opt.AddAllowedClass(typeof(DBText), true);
            opt.AddAllowedClass(typeof(MText), true);
            var res = ed.GetEntity(opt);
            if (res.Status == PromptStatus.OK)
            {
                return res.ObjectId;
            }
            else
            {
                return ObjectId.Null;
            }
        }
    }
}
