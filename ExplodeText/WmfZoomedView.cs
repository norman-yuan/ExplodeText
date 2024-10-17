using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExplodeText
{
    public class WmfZoomedView : IDisposable
    {
        private readonly Editor _ed;

        private ViewTableRecord _currentView = null; 

        public WmfZoomedView(Editor ed, Extents3d zoomExtents)
        {
            _ed = ed;
            _currentView = _ed.GetCurrentView();
            ZoomToExtents(zoomExtents);
        }

        public void Dispose()
        {
            if (_currentView != null)
            {
                _ed.SetCurrentView(_currentView);
            }
        }

        private void ZoomToExtents(Extents3d zoomExtents)
        {
            var d = (zoomExtents.MaxPoint.X - zoomExtents.MinPoint.X) / 50;

            var pt1 = new[] 
            { 
                zoomExtents.MinPoint.X - d, 
                zoomExtents.MinPoint.Y - d, 
                zoomExtents.MinPoint.Z 
            };
            var pt2 = new[] 
            { 
                zoomExtents.MaxPoint.X + d, 
                zoomExtents.MaxPoint.Y + d, 
                zoomExtents.MaxPoint.Z 
            };

            dynamic comApp = Application.AcadApplication;
            comApp.ZoomWindow(pt1, pt2);
        }

    }
}
