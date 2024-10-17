using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System.Collections.Generic;
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
                List<ObjectId> explodedCurves;
                var exploder = new TextExploder(dwg);
                if (txtId.ObjectClass.DxfName.ToUpper() == "TEXT")
                {
                    explodedCurves = exploder.ExplodeDBText(txtId);
                }
                else
                {
                    explodedCurves = exploder.ExplodeMText(txtId);
                }

                // Do something with explosion-generated curves
                editor.WriteMessage(
                    $"\n{explodedCurves.Count} curves generated from text explodion.\n");
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
            opt.SetRejectMessage("\nInvalid: must be a TEXT/MTEXT entity.");
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
