using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System.Drawing;
using System.Reflection.Metadata;
using System.Security.Cryptography;
using System.Transactions;

namespace RevitWallConnectivityChecker
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class WallConnectivityChecker : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Получение текущего документа Revit
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Получение всех стен в проекте
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(Wall));
            var walls = collector.ToElements();

            // Проверка связности стен и окрашивание неприсоединенных стен
            foreach (Wall wall in walls)
            {
                if (!IsWallConnected(wall))
                {
                    // Неприсоединенная стена
                    ChangeWallColor(doc, wall, Color.Red);
                }
            }

            // Перерисовка представления для обновления цвета стен на экране
            uidoc.RefreshActiveView();

            return Result.Succeeded;
        }

        private bool IsWallConnected(Wall wall)
        {
            // Получение контура стены
            LocationCurve locationCurve = wall.Location as LocationCurve;
            Curve curve = locationCurve.Curve;

            // Поиск других стен, которые пересекают контур данной стены
            FilteredElementCollector collector = new FilteredElementCollector(wall.Document);
            collector.OfClass(typeof(Wall));
            collector.WherePasses(new BoundingBoxIntersectsFilter(curve.GetBoundingBox()));

            foreach (Wall intersectingWall in collector)
            {
                if (intersectingWall.Id != wall.Id)
                {
                    return true;
                }
            }

            return false;
        }

        private void ChangeWallColor(Document doc, Wall wall, Color color)
        {
            // Получение графических элементов стены
            GraphicsStyle graphicsStyle = wall.GetGraphicsStyle(GraphicsStyleType.Projection);
            ElementId graphicsStyleId = graphicsStyle.Id;

            // Создание графических настроек для изменения цвета
            OverrideGraphicSettings settings = new OverrideGraphicSettings();
            settings.SetProjectionLineColor(color);

            // Применение графических настроек к стене
            using (Transaction tx = new Transaction(doc, "Change Wall Color"))
            {
                tx.Start();

                // Применение графических настроек к стене
                doc.ActiveView.SetElementOverrides(wall.Id, settings);

                // Обновление графических настроек для графического стиля стены
                doc.ActiveView.GetElementOverrides(wall.Id).SetProjectionLinePatternId(graphicsStyleId);

                tx.Commit();
            }
        }
    }
}
