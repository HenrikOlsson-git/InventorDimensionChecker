using Inventor;

namespace DimensionCheckerPlugin
{
    public class DimensionChecker
    {
        private Inventor.Application _inventorApp;

        public void Activate(Inventor.Application inventorApp)
        {
            _inventorApp = inventorApp;
        }

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
            PartComponentDefinition compDef = partDoc.ComponentDefinition as PartComponentDefinition;

            if (compDef != null)
            {
                var parameters = compDef.Parameters;
                foreach (UserParameter param in parameters.UserParameters)
                {
                    _inventorApp.StatusBarText = $"Model Parameter: {param.Name} = {param.Expression}";
                }
            }
            else
            {
                _inventorApp.StatusBarText = "Could not access parameters for this document.";
            }
        }

        private void CompareDrawingDimensions(DrawingDocument drawingDoc)
        {
            foreach (Sheet sheet in drawingDoc.Sheets)
            {
                foreach (DrawingDimension dim in sheet.DrawingDimensions)
                {
                    _inventorApp.StatusBarText = $"Drawing Dimension: {dim.Text}";
                }
            }
        }
    }
}