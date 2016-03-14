using System;
using System.Windows;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.Utils;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ResizeLab
{
    public partial class ResizeLabPaneWPF
    {
        private bool _isPreview = false;
        public void ShowErrorMessageBox(string content, Exception exception = null)
        {
            if (exception != null)
            {
                Views.ErrorDialogWrapper.ShowDialog("Error", content, exception);
            }
            else
            {
                MessageBox.Show(content, "Error");
            }
        }


        public void Preview(PowerPoint.ShapeRange selectedShapes, Action<PowerPoint.ShapeRange> previewAction, int minNumberofSelectedShapes)
        {
            if (selectedShapes == null || selectedShapes.Count < minNumberofSelectedShapes) return;

            this.StartNewUndoEntry();
            previewAction.Invoke(selectedShapes);
            _isPreview = true;
        }

        public void Preview(PowerPoint.ShapeRange selectedShapes, float referenceWidth, float referenceHeight, Action<PowerPoint.ShapeRange, float, float, bool> previewAction)
        {
            if (selectedShapes == null) return;

            this.StartNewUndoEntry();
            previewAction.Invoke(selectedShapes, referenceWidth, referenceHeight, IsAspectRatioLocked);
            _isPreview = true;
        }

        public void Reset(int numberOfUndo = 1)
        {
            //var selectedShapes = GetSelectedShapes(false);

            if (!_isPreview) return;

            for (int i = 0; i < numberOfUndo; i++)
            {
                this.ExecuteOfficeCommand("Undo");
            }
            _isPreview = false;
            GC.Collect();
        }

        public void ExecuteResizeAction(PowerPoint.ShapeRange selectedShapes, Action<PowerPoint.ShapeRange> resizeAction)
        {
            if (selectedShapes == null) return;

            Reset();
            this.StartNewUndoEntry();
            resizeAction.Invoke(selectedShapes);
        }

        public void ExecuteResizeAction(PowerPoint.ShapeRange selectedShapes, float slideWidth, float slideHeight, Action<PowerPoint.ShapeRange, float, float, bool> resizeAction)
        {
            if (selectedShapes == null) return;

            Reset();
            this.StartNewUndoEntry();
            resizeAction.Invoke(selectedShapes, slideWidth, slideHeight, IsAspectRatioLocked);
        }
    }
}
