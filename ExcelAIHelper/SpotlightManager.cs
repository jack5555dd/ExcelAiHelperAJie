using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System;

namespace ExcelAIHelper
{
    internal static class SpotlightManager
    {
        private const string MASK_NAME = "AI_SpotlightMask";
        private const string SPOTLIGHT_GROUP = "AI_SpotlightGroup";

        public static void Toggle()
        {
            var app = Globals.ThisAddIn.Application;
            Excel.Worksheet ws = app.ActiveSheet;

            Excel.Shape existingGroup;
            try { existingGroup = ws.Shapes.Item(SPOTLIGHT_GROUP); }
            catch { existingGroup = null; }

            if (existingGroup == null)
                Apply(ws, app.Selection as Excel.Range);
            else
                Remove(ws);
        }

        private static void Apply(Excel.Worksheet ws, Excel.Range sel)
        {
            try
            {
                // Get worksheet dimensions (much larger than visible area to cover scrolling)
                double wsWidth = 2000;  // Large enough to cover most worksheets
                double wsHeight = 1500;
                
                // Get selection position and size
                double selLeft = sel.Left;
                double selTop = sel.Top;
                double selWidth = sel.Width;
                double selHeight = sel.Height;
                
                // Create 4 rectangles around the selection to form a "frame" effect
                // Top rectangle
                var topMask = ws.Shapes.AddShape(
                    Office.MsoAutoShapeType.msoShapeRectangle,
                    0, 0, (float)wsWidth, (float)selTop);
                ConfigureMaskShape(topMask, "AI_SpotlightTop");
                
                // Bottom rectangle  
                var bottomMask = ws.Shapes.AddShape(
                    Office.MsoAutoShapeType.msoShapeRectangle,
                    0, (float)(selTop + selHeight), (float)wsWidth, (float)(wsHeight - selTop - selHeight));
                ConfigureMaskShape(bottomMask, "AI_SpotlightBottom");
                
                // Left rectangle
                var leftMask = ws.Shapes.AddShape(
                    Office.MsoAutoShapeType.msoShapeRectangle,
                    0, (float)selTop, (float)selLeft, (float)selHeight);
                ConfigureMaskShape(leftMask, "AI_SpotlightLeft");
                
                // Right rectangle
                var rightMask = ws.Shapes.AddShape(
                    Office.MsoAutoShapeType.msoShapeRectangle,
                    (float)(selLeft + selWidth), (float)selTop, (float)(wsWidth - selLeft - selWidth), (float)selHeight);
                ConfigureMaskShape(rightMask, "AI_SpotlightRight");
                
                // Add highlight border to selection
                sel.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                sel.Borders.Color = 0x0000FF;  // Red border for spotlight effect
                sel.Borders.Weight = Excel.XlBorderWeight.xlThick;
                
                // Group all shapes together for easy management
                try
                {
                    var shapeNames = new string[] { "AI_SpotlightTop", "AI_SpotlightBottom", "AI_SpotlightLeft", "AI_SpotlightRight" };
                    var shapeRange = ws.Shapes.get_Range(shapeNames);
                    var group = shapeRange.Group();
                    group.Name = SPOTLIGHT_GROUP;
                }
                catch { }
            }
            catch (Exception ex)
            {
                // Fallback: just highlight the selection
                try
                {
                    sel.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    sel.Borders.Color = 0x0000FF;
                    sel.Borders.Weight = Excel.XlBorderWeight.xlThick;
                }
                catch { }
            }
        }
        
        private static void ConfigureMaskShape(Excel.Shape shape, string name)
        {
            shape.Name = name;
            shape.Fill.ForeColor.RGB = 0x000000;  // Black mask
            shape.Fill.Transparency = 0.6f;       // 60% transparent
            shape.Line.Visible = Office.MsoTriState.msoFalse;
            shape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
        }

        private static void Remove(Excel.Worksheet ws)
        {
            try
            {
                // Remove the spotlight group if it exists
                try
                {
                    var group = ws.Shapes.Item(SPOTLIGHT_GROUP);
                    group.Delete();
                }
                catch { }
                
                // Remove individual spotlight shapes if they exist (fallback)
                string[] shapeNames = { "AI_SpotlightTop", "AI_SpotlightBottom", "AI_SpotlightLeft", "AI_SpotlightRight" };
                foreach (string shapeName in shapeNames)
                {
                    try
                    {
                        var shape = ws.Shapes.Item(shapeName);
                        shape.Delete();
                    }
                    catch { }
                }
                
                // Clear selection borders
                try
                {
                    var selection = ws.Application.Selection as Excel.Range;
                    if (selection != null)
                    {
                        selection.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                    }
                }
                catch { }
            }
            catch { }
        }

        public static void RemoveIfExists()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                Excel.Worksheet ws = app.ActiveSheet;
                
                Remove(ws);
            }
            catch { }
        }
    }
}