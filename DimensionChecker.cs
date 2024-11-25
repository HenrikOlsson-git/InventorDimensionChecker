using System;
using Inventor;

namespace DimensionCheckerPlugin
{
    public class DimensionChecker : ApplicationAddInServer
    {
        private Inventor.Application _inventorApp;

        // Entry point for the plugin
        public void Activate(ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            _inventorApp = addInSiteObject.Application;
        }

        public void Deactivate()
        {
            _inventorApp = null;
        }

        public void ExecuteCommand(int commandID) { }

        public object Automation => null;

        // Main logic for dimension comparison
        public void CompareDimensions()
        {
            // Ensure an active document is open
            Document activeDoc = _inventorApp.ActiveDocument;
            if (activeDoc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                CompareModelDimensions((PartDocument)activeDoc);
            }
            else if (activeDoc.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
            {
                CompareDrawingDimensions((DrawingDocument)activeDoc);
            }
            else
            {
                _inventorApp.StatusBarText = "Open a part or drawing document to use this tool.";
            }
        }

        private void CompareModelDimensions(PartDocument partDoc)
        {
            ComponentDefinition compDef = partDoc.ComponentDefinition;

            // Extract all parameters from the model
            var parameters = compDef.Parameters;
            foreach (UserParameter param in parameters.UserParameters)
            {
                Console.WriteLine($"Model Parameter: {param.Name} = {param.Expression}");
            }
        }

        private void CompareDrawingDimensions(DrawingDocument drawingDoc)
        {
            foreach (Sheet sheet in drawingDoc.Sheets)
            {
                foreach (DrawingView view in sheet.DrawingViews)
                {
                    foreach (DrawingDimension dim in view.DrawingDimensions)
                    {
                        Console.WriteLine($"Drawing Dimension: {dim.Text}");
                    }
                }
            }
        }
    }
}

