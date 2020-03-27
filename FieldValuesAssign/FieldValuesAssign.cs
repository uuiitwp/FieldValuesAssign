using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace FieldValuesAssign
{
    public class FieldValuesAssign : ESRI.ArcGIS.Desktop.AddIns.Button
    {
        public FieldValuesAssign()
        {
        }

        protected override void OnClick()
        {
            MainWindow mw = new MainWindow();
            mw.Show();
            ArcMap.Application.CurrentTool = null;
        }
        protected override void OnUpdate()
        {
            Enabled = ArcMap.Application != null;
        }
    }

}
