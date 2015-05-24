using System;
using System.Globalization;
using Microsoft.Office.Interop.Visio;

namespace GedToVisio.Visio
{
    public static class VisioHelper
    {
        /// <summary>
        /// Draw Rectangle with 2 connections points.
        /// </summary>
        public static Shape DrawRectangle(Page page, double x1, double y1, double x2, double y2)
        {
            Shape rect = page.DrawRectangle(x1, y1, x2, y2);

            // add connection point1
            short rowIndex1 = rect.AddRow(
                (short)Microsoft.Office.Interop.Visio.VisSectionIndices.visSectionConnectionPts,
                (short)Microsoft.Office.Interop.Visio.VisRowIndices.visRowLast,
                (short)Microsoft.Office.Interop.Visio.VisRowTags.visTagCnnctPt);

            Row row1 = rect.Section[(short)Microsoft.Office.Interop.Visio.VisSectionIndices.visSectionConnectionPts][rowIndex1];
            row1.CellU[(short)Microsoft.Office.Interop.Visio.VisCellIndices.visCnnctX].FormulaU = "Width*0";
            row1.CellU[(short)Microsoft.Office.Interop.Visio.VisCellIndices.visCnnctY].FormulaU = "Height*0.5";
            row1.CellU[(short)Microsoft.Office.Interop.Visio.VisCellIndices.visCnnctDirX].FormulaU = "1";
            row1.CellU[(short)Microsoft.Office.Interop.Visio.VisCellIndices.visCnnctDirY].FormulaU = "0";
            row1.CellU[(short)Microsoft.Office.Interop.Visio.VisCellIndices.visCnnctType].FormulaU =
                ((short)Microsoft.Office.Interop.Visio.tagVisCellVals.visCnnctTypeInward).ToString(CultureInfo.InvariantCulture);

            // add connection point2
            short rowIndex2 = rect.AddRow(
                (short)Microsoft.Office.Interop.Visio.VisSectionIndices.visSectionConnectionPts,
                (short)Microsoft.Office.Interop.Visio.VisRowIndices.visRowLast,
                (short)Microsoft.Office.Interop.Visio.VisRowTags.visTagCnnctPt);

            Row row2 = rect.Section[(short)Microsoft.Office.Interop.Visio.VisSectionIndices.visSectionConnectionPts][rowIndex2];
            row2.CellU[(short)Microsoft.Office.Interop.Visio.VisCellIndices.visCnnctX].FormulaU = "Width*1";
            row2.CellU[(short)Microsoft.Office.Interop.Visio.VisCellIndices.visCnnctY].FormulaU = "Height*0.5";
            row2.CellU[(short)Microsoft.Office.Interop.Visio.VisCellIndices.visCnnctDirX].FormulaU = "-1";
            row2.CellU[(short)Microsoft.Office.Interop.Visio.VisCellIndices.visCnnctDirY].FormulaU = "0";
            row2.CellU[(short)Microsoft.Office.Interop.Visio.VisCellIndices.visCnnctType].FormulaU =
                ((short)Microsoft.Office.Interop.Visio.tagVisCellVals.visCnnctTypeInward).ToString(CultureInfo.InvariantCulture);

            return rect;
        }


        /// <summary>
        /// Соединяет фигуры через указанные Connection Points.
        /// Если ConnectionPointIndex null, соединяет сами фигуры, линией к контуру фигуры.
        /// ConnectionPointIndex начинается с 0.
        /// </summary>
        public static Shape Connect(Shape shapeFrom, short? shapeFromConnectionPointIndex, Shape shapeTo, short? shapeToConnectionPointIndex)
        {
            // Verify that the shapes are on the same page.
            if (shapeFrom.ContainingPage == null || shapeTo.ContainingPage == null ||
                !shapeFrom.ContainingPage.Equals(shapeTo.ContainingPage))
            {
                throw new ArgumentException("The shapes are not on the same page.");
            }

            // Get the Application object from the shape.
            var application = shapeFrom.Application;

            // Access the Basic Flowchart Shapes stencil from the documents collection of the application.
            var stencil = application.Documents.OpenEx("Basic Flowchart Shapes (US units).vss",
                (short)Microsoft.Office.Interop.Visio.VisOpenSaveArgs.visOpenDocked);

            // Get the dynamic connector master on the stencil by its universal name.
            var masterInStencil = stencil.Masters.ItemU["Dynamic Connector"];

            // Drop the dynamic connector on the active page.
            var connector = application.ActivePage.Drop(masterInStencil, 0, 0);

            var beginXCell = connector.CellsU["BeginX"];
            if (shapeFromConnectionPointIndex == null)
            {
                // Glue connector to shape with dynamic glue
                beginXCell.GlueTo(shapeFrom.CellsSRC[
                    (short)Microsoft.Office.Interop.Visio.VisSectionIndices.visSectionObject,
                    (short)Microsoft.Office.Interop.Visio.VisRowIndices.visRowXFormOut,
                    (short)Microsoft.Office.Interop.Visio.VisCellIndices.visXFormPinX]);
            }
            else
            {
                // Glue connector to connection point
                beginXCell.GlueTo(shapeFrom.CellsSRC[
                    (short)Microsoft.Office.Interop.Visio.VisSectionIndices.visSectionConnectionPts,
                    (short)shapeFromConnectionPointIndex,
                    0]);
            }


            var endXCell = connector.CellsU["EndX"];
            if (shapeToConnectionPointIndex == null)
            {
                // Glue connector to shape with dynamic glue
                endXCell.GlueTo(shapeTo.CellsSRC[
                    (short)Microsoft.Office.Interop.Visio.VisSectionIndices.visSectionObject,
                    (short)Microsoft.Office.Interop.Visio.VisRowIndices.visRowXFormOut,
                    (short)Microsoft.Office.Interop.Visio.VisCellIndices.visXFormPinX]);
            }
            else
            {
                // Glue connector to connection point
                endXCell.GlueTo(shapeTo.CellsSRC[
                    (short)Microsoft.Office.Interop.Visio.VisSectionIndices.visSectionConnectionPts,
                    (short)shapeToConnectionPointIndex,
                    0]);
            }


            // прямая соединительная линия
            connector.CellsU["ShapeRouteStyle"].FormulaU = "16";

            return connector;
        }


        public static string GetCustomProperty(Shape shape, string propertyName)
        {
            var propName = "Prop." + propertyName;

            var customPropertyCell = shape.CellsU[propName];
            return customPropertyCell.ResultStr[VisUnitCodes.visNoCast];
        }


        public static void SetCustomProperty(Shape shape, string propertyName, string value) 
        {
            SetCustomProperty(shape, propertyName, propertyName, value);
        }


        public static void SetCustomProperty(Shape shape, string propertyName, string label, string value)
        {
            short rowIndex = shape.AddNamedRow((short)VisSectionIndices.visSectionProp, propertyName,
                                               (short)VisRowIndices.visRowProp);

            shape.CellsSRC[(short)VisSectionIndices.visSectionProp, rowIndex, (short)VisCellIndices.visCustPropsLabel].
                FormulaU = StringToFormulaForString(label);

            shape.CellsSRC[(short)VisSectionIndices.visSectionProp, rowIndex, (short)VisCellIndices.visCustPropsType].FormulaU
                =
                StringToFormulaForString(
                    ((short)VisCellVals.visPropTypeString).ToString(System.Globalization.CultureInfo.InvariantCulture));

            shape.CellsSRC[(short)VisSectionIndices.visSectionProp, rowIndex, (short)VisCellIndices.visCustPropsValue].
                FormulaU = StringToFormulaForString(value);
        }

        public static string StringToFormulaForString(string inputValue)
        {
            const string quote = "\"";
            const string quoteQuote = "\"\"";

            return quote + (inputValue ?? String.Empty).Replace(quote, quoteQuote) + quote;
        }

    }
}
