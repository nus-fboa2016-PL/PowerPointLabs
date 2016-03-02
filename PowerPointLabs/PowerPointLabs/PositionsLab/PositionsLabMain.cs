using System;
using System.Collections.Generic;
using Microsoft.Office.Core;
using PowerPointLabs.Utils;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using AutoShape = Microsoft.Office.Core.MsoAutoShapeType;
using System.Diagnostics;
using Drawing = System.Drawing;
using System.IO;

namespace PowerPointLabs.PositionsLab
{
    class PositionsLabMain
    {
        private static bool _isInit;
        private const float Epsilon = 0.00001f;
        private const float RotateLeft = 90f;
        private const float RotateRight = 270f;
        private const float RotateUp = 0f;
        private const float RotateDown = 180f;
        private const int None = -1;
        private const int Right = 0;
        private const int Down = 1;
        private const int Left = 2;
        private const int Up = 3;
        private const int Leftorright = 4;
        private const int Upordown = 5;

        //For Grid
        public enum GridAlignment
        {
            AlignLeft,
            AlignCenter,
            AlignRight
        }

        private static GridAlignment _distributeGridAlignment;
        private static float _marginTop = 5;
        private static float _marginBottom = 5;
        private static float _marginLeft = 5;
        private static float _marginRight = 5;

        private static Dictionary<MsoAutoShapeType, float> shapeDefaultUpAngle;
        private static bool _alignUseSlideAsReference = false;
        private static bool _distributeUseSlideAsReference = false;

        private static readonly ISet<MsoAutoShapeType> ShapeFillsBoundingBox = new HashSet<MsoAutoShapeType>
        {
            MsoAutoShapeType.msoShapeRectangle,
            MsoAutoShapeType.msoShapeBevel,
            MsoAutoShapeType.msoShapeFrame,
            MsoAutoShapeType.msoShapeFlowchartProcess,
            MsoAutoShapeType.msoShapeFlowchartPredefinedProcess,
            MsoAutoShapeType.msoShapeFlowchartInternalStorage,
            MsoAutoShapeType.msoShapeActionButtonBackorPrevious,
            MsoAutoShapeType.msoShapeActionButtonForwardorNext,
            MsoAutoShapeType.msoShapeActionButtonBeginning,
            MsoAutoShapeType.msoShapeActionButtonEnd,
            MsoAutoShapeType.msoShapeActionButtonHome,
            MsoAutoShapeType.msoShapeActionButtonInformation,
            MsoAutoShapeType.msoShapeActionButtonReturn,
            MsoAutoShapeType.msoShapeActionButtonMovie,
            MsoAutoShapeType.msoShapeActionButtonDocument,
            MsoAutoShapeType.msoShapeActionButtonSound,
            MsoAutoShapeType.msoShapeActionButtonHelp,
            MsoAutoShapeType.msoShapeActionButtonCustom
        }; 

        private static readonly ISet<MsoAutoShapeType> ShapeHasNoNodes = new HashSet<MsoAutoShapeType>()
        {
            MsoAutoShapeType.msoShapeOval
        }; 

        private static readonly ISet<MsoAutoShapeType> ShapeNotSupportedForInternalVertices = new HashSet<MsoAutoShapeType>
        {
            MsoAutoShapeType.msoShapeChord,
            MsoAutoShapeType.msoShapeHeart
        }; 

        #region API

        #region Class Methods

        /// <summary>
        /// Tells the Positions Lab to use the slide as the reference point for Align methods
        /// </summary>
        public static void AlignReferToSlide()
        {
            _alignUseSlideAsReference = true;
        }

        /// <summary>
        /// Tells the Positions Lab to use first selected shape as reference shape for Align methods
        /// </summary>
        public static void AlignReferToShape()
        {
            _alignUseSlideAsReference = false;
        }

        /// <summary>
        /// Tells the Position Lab to use the slide as the reference point for Distribute methods
        /// </summary>
        public static void DistributeReferToSlide()
        {
            _distributeUseSlideAsReference = true;
        }

        /// <summary>
        /// Tells the Positions Lab to use first selected shape as reference shape for Distribute methods
        /// </summary>
        public static void DistributeReferToShape()
        {
            _distributeUseSlideAsReference = false;
        }

        public static void SetDistributeGridAlignment(GridAlignment alignment)
        {
            _distributeGridAlignment = alignment;
        }

        public static void SetDistributeMarginTop (float marginTop)
        {
            _marginTop = marginTop;
        }

        public static void SetDistributeMarginBottom(float marginBottom)
        {
            _marginBottom = marginBottom;
        }

        public static void SetDistributeMarginLeft(float marginLeft)
        {
            _marginLeft = marginLeft;
        }

        public static void SetDistributeMarginRight(float marginRight)
        {
            _marginRight = marginRight;
        }

        #endregion

        #region Align
        public static void AlignLeft(List<Shape> selectedShapes)
        {
            if (_alignUseSlideAsReference)
            {
                foreach (var s in selectedShapes)
                {
                    var allPointsOfShape = Graphics.GetRealCoordinates(s);
                    var leftMost = Graphics.LeftMostPoint(allPointsOfShape);
                    s.IncrementLeft(-leftMost.X);
                }
            }
            else
            {
                if (selectedShapes.Count < 2)
                {
                    //Error
                    return;
                }

                var refShape = selectedShapes[0];
                var allPointsOfRef = Graphics.GetRealCoordinates(refShape);
                var leftMostRef = Graphics.LeftMostPoint(allPointsOfRef);

                for (var i = 1; i < selectedShapes.Count; i++)
                {
                    var s = selectedShapes[i];
                    var allPoints = Graphics.GetRealCoordinates(s);
                    var leftMost = Graphics.LeftMostPoint(allPoints);
                    s.IncrementLeft(leftMostRef.X - leftMost.X);
                }
            }
        }

        public static void AlignRight(List<Shape> selectedShapes, float slideWidth)
        {
            if (_alignUseSlideAsReference)
            {
                foreach (var s in selectedShapes)
                {
                    var allPointsOfShape = Graphics.GetRealCoordinates(s);
                    var leftMost = Graphics.LeftMostPoint(allPointsOfShape);
                    var shapeWidth = Graphics.RealWidth(allPointsOfShape);
                    s.IncrementLeft(slideWidth - leftMost.X - shapeWidth);
                }
            }
            else
            {
                if (selectedShapes.Count < 2)
                {
                    //Error
                    return;
                }

                var refShape = selectedShapes[0];
                var allPointsOfRef = Graphics.GetRealCoordinates(refShape);
                var rightMostRef = Graphics.RightMostPoint(allPointsOfRef);

                for (var i = 1; i < selectedShapes.Count; i++)
                {
                    var s = selectedShapes[i];
                    var allPoints = Graphics.GetRealCoordinates(s);
                    var rightMost = Graphics.RightMostPoint(allPoints);
                    s.IncrementLeft(rightMostRef.X - rightMost.X);
                }
            }
        }

        public static void AlignTop(List<Shape> selectedShapes)
        {
            if (_alignUseSlideAsReference)
            {
                foreach (var s in selectedShapes)
                {
                    var allPointsOfShape = Graphics.GetRealCoordinates(s);
                    var topMost = Graphics.TopMostPoint(allPointsOfShape);
                    s.IncrementTop(-topMost.Y);
                }
            }
            else
            {
                if (selectedShapes.Count < 2)
                {
                    //Error
                    return;
                }

                var refShape = selectedShapes[0];
                var allPointsOfRef = Graphics.GetRealCoordinates(refShape);
                var topMostRef = Graphics.TopMostPoint(allPointsOfRef);

                for (var i = 1; i < selectedShapes.Count; i++)
                {
                    var s = selectedShapes[i];
                    var allPoints = Graphics.GetRealCoordinates(s);
                    var topMost = Graphics.TopMostPoint(allPoints);
                    s.IncrementTop(topMostRef.Y - topMost.Y);
                }
            }
        }

        public static void AlignBottom(List<Shape> selectedShapes, float slideHeight)
        {
            if (_alignUseSlideAsReference)
            {
                foreach (var s in selectedShapes)
                {
                    var allPointsOfShape = Graphics.GetRealCoordinates(s);
                    var topMost = Graphics.TopMostPoint(allPointsOfShape);
                    var shapeHeight = Graphics.RealHeight(allPointsOfShape);
                    s.IncrementTop(slideHeight - topMost.Y - shapeHeight);
                }
            }
            else
            {
                if (selectedShapes.Count < 2)
                {
                    //Error
                    return;
                }

                var refShape = selectedShapes[0];
                var allPointsOfRef = Graphics.GetRealCoordinates(refShape);
                var lowestRef = Graphics.BottomMostPoint(allPointsOfRef);

                for (var i = 1; i < selectedShapes.Count; i++)
                {
                    var s = selectedShapes[i];
                    var allPoints = Graphics.GetRealCoordinates(s);
                    var lowest = Graphics.BottomMostPoint(allPoints);
                    s.IncrementTop(lowestRef.Y - lowest.Y);
                }
            }
        }

        public static void AlignMiddle(List<Shape> selectedShapes, float slideHeight)
        {
            if (_alignUseSlideAsReference)
            {
                foreach (var s in selectedShapes)
                {
                    var allPointsOfShape = Graphics.GetRealCoordinates(s);
                    var topMost = Graphics.TopMostPoint(allPointsOfShape);
                    var shapeHeight = Graphics.RealHeight(allPointsOfShape);
                    s.IncrementTop(slideHeight/2 - topMost.Y - shapeHeight/2);
                }
            }
            else
            {
                if (selectedShapes.Count < 2)
                {
                    //Error
                    return;
                }

                var refShape = selectedShapes[0];
                var originRef = Graphics.GetCenterPoint(refShape);

                for (var i = 1; i < selectedShapes.Count; i++)
                {
                    var s = selectedShapes[i];
                    var origin = Graphics.GetCenterPoint(s);
                    s.IncrementTop(originRef.Y - origin.Y);
                }
            }
        }

        public static void AlignCenter(List<Shape> selectedShapes, float slideWidth, float slideHeight)
        {
            if (_alignUseSlideAsReference)
            {
                foreach (var s in selectedShapes)
                {
                    var allPointsOfShape = Graphics.GetRealCoordinates(s);
                    var topMost = Graphics.TopMostPoint(allPointsOfShape);
                    var leftMost = Graphics.LeftMostPoint(allPointsOfShape);
                    var shapeHeight = Graphics.RealHeight(allPointsOfShape);
                    var shapeWidth = Graphics.RealWidth(allPointsOfShape);
                    s.IncrementTop(slideHeight/2 - topMost.Y - shapeHeight/2);
                    s.IncrementLeft(slideWidth/2 - leftMost.X - shapeWidth/2);
                }
            }
            else
            {
                if (selectedShapes.Count < 2)
                {
                    //Error
                    return;
                }

                var refShape = selectedShapes[0];
                var originRef = Graphics.GetCenterPoint(refShape);

                for (var i = 1; i < selectedShapes.Count; i++)
                {
                    var s = selectedShapes[i];
                    var origin = Graphics.GetCenterPoint(s);
                    s.IncrementLeft(originRef.X - origin.X);
                    s.IncrementTop(originRef.Y - origin.Y);
                }
            }
        }

        #endregion

        #region Snap
        public static void SnapVertical(List<Shape> selectedShapes)
        {
            if (!_isInit)
            {
                Init();
                _isInit = true;
            }

            foreach (var s in selectedShapes)
            {
                SnapShapeVertical(s);
            }
        }

        public static void SnapHorizontal(List<Shape> selectedShapes)
        {
            if (!_isInit)
            {
                Init();
                _isInit = true;
            }

            foreach (var s in selectedShapes)
            {
                SnapShapeHorizontal(s);
            }
        }

        public static void SnapAway(List<Shape> shapes)
        {
            if (!_isInit)
            {
                Init();
                _isInit = true;
            }

            if (shapes.Count <= 1)
            {
                return;
            }

            var refShapeCenter = Graphics.GetCenterPoint(shapes[0]);
            var isAllSameDir = true;
            var lastDir = -1;

            for (var i = 1; i < shapes.Count; i++)
            {
                var shape = shapes[i];
                var shapeCenter = Graphics.GetCenterPoint(shape);
                var angle = (float)AngleBetweenTwoPoints(refShapeCenter, shapeCenter);

                var dir = GetDirectionWrtRefShape(shape, angle);

                if (i == 1)
                {
                    lastDir = dir;
                }

                if (!IsSameDirection(lastDir, dir))
                {
                    isAllSameDir = false;
                    break;
                }

                //only maintain in one direction instead of dual direction
                if (dir < Leftorright)
                {
                    lastDir = dir; 
                }
            }

            if (!isAllSameDir || lastDir == None)
            {
                lastDir = 0;
            }
            else
            {
                lastDir++;
            }

            for (var i = 1; i < shapes.Count; i++)
            {
                var shape = shapes[i];
                var shapeCenter = Graphics.GetCenterPoint(shape);
                var angle = (float) AngleBetweenTwoPoints(refShapeCenter, shapeCenter);

                float defaultUpAngle = 0;
                var hasDefaultDirection = shapeDefaultUpAngle.TryGetValue(shape.AutoShapeType, out defaultUpAngle);

                if (hasDefaultDirection)
                {
                    shape.Rotation = (defaultUpAngle + angle) + lastDir * 90;
                }
                else
                {
                    if (IsVertical(shape))
                    {
                        shape.Rotation = angle + lastDir * 90;
                    }
                    else
                    {
                        shape.Rotation = (angle - 90) + lastDir * 90;
                    }
                }
            }
        }

        public static void SnapShapeVertical(Shape shape)
        {
            if (IsVertical(shape))
            {
                SnapTo0Or180(shape);
            }
            else
            {
                SnapTo90Or270(shape);
            }
        }

        public static void SnapShapeHorizontal(Shape shape)
        {
            if (IsVertical(shape))
            {
                SnapTo90Or270(shape);
            }
            else
            {
                SnapTo0Or180(shape);
            }
        }

        private static void SnapTo0Or180 (Shape shape)
        {
            var rotation = shape.Rotation;

            if (rotation >= 90 && rotation < 270)
            {
                shape.Rotation = 180;
            }
            else
            {
                shape.Rotation = 0;
            }
        }

        private static void SnapTo90Or270(Shape shape)
        {
            var rotation = shape.Rotation;

            if (rotation >= 0 && rotation < 180)
            {
                shape.Rotation = 90;
            }
            else
            {
                shape.Rotation = 270;
            }
        }

        private static bool IsVertical(Shape shape)
        {
            return shape.Height > shape.Width;
        }
        #endregion

        #region Adjoin
        public static void AdjoinHorizontal(List<Shape> selectedShapes)
        {
            if (selectedShapes.Count < 2)
            {
                //Error
                return;
            }

            var refShape = selectedShapes[0];
            var sortedShapes = Graphics.SortShapesByLeft(selectedShapes);
            var refShapeIndex = sortedShapes.IndexOf(refShape);

            var allPointsOfRef = Graphics.GetRealCoordinates(refShape);
            var centerOfRef = Graphics.GetCenterPoint(refShape);

            var mostLeft = Graphics.LeftMostPoint(allPointsOfRef).X;
            //For all shapes left of refShape, adjoin them from closest to refShape
            for (var i = refShapeIndex - 1; i >= 0; i--)
            {
                var neighbour = sortedShapes[i];
                var allPointsOfNeighbour = Graphics.GetRealCoordinates(neighbour);
                var rightOfShape = Graphics.RightMostPoint(allPointsOfNeighbour).X;
                neighbour.IncrementLeft(mostLeft - rightOfShape);
                neighbour.IncrementTop(centerOfRef.Y - Graphics.GetCenterPoint(neighbour).Y);

                mostLeft = Graphics.LeftMostPoint(allPointsOfNeighbour).X + mostLeft - rightOfShape;
            }

            var mostRight = Graphics.RightMostPoint(allPointsOfRef).X;
            //For all shapes right of refShape, adjoin them from closest to refShape
            for (var i = refShapeIndex + 1; i < sortedShapes.Count; i++)
            {
                var neighbour = sortedShapes[i];
                var allPointsOfNeighbour = Graphics.GetRealCoordinates(neighbour);
                var leftOfShape = Graphics.LeftMostPoint(allPointsOfNeighbour).X;
                neighbour.IncrementLeft(mostRight - leftOfShape);
                neighbour.IncrementTop(centerOfRef.Y - Graphics.GetCenterPoint(neighbour).Y);

                mostRight = Graphics.RightMostPoint(allPointsOfNeighbour).X + mostRight - leftOfShape;
            }
        }

        public static void AdjoinVertical(List<Shape> selectedShapes)
        {
            if (selectedShapes.Count < 2)
            {
                //Error
                return;
            }

            var refShape = selectedShapes[0];
            var sortedShapes = Graphics.SortShapesByTop(selectedShapes);
            var refShapeIndex = sortedShapes.IndexOf(refShape);

            var allPointsOfRef = Graphics.GetRealCoordinates(refShape);
            var centerOfRef = Graphics.GetCenterPoint(refShape);

            var mostTop = Graphics.TopMostPoint(allPointsOfRef).Y;
            //For all shapes above refShape, adjoin them from closest to refShape
            for (var i = refShapeIndex - 1; i >= 0; i--)
            {
                var neighbour = sortedShapes[i];
                var allPointsOfNeighbour = Graphics.GetRealCoordinates(neighbour);
                var bottomOfShape = Graphics.BottomMostPoint(allPointsOfNeighbour).Y;
                neighbour.IncrementLeft(centerOfRef.X - Graphics.GetCenterPoint(neighbour).X);
                neighbour.IncrementTop(mostTop - bottomOfShape);

                mostTop = Graphics.TopMostPoint(allPointsOfNeighbour).Y + mostTop - bottomOfShape;
            }

            var lowest = Graphics.BottomMostPoint(allPointsOfRef).Y;
            //For all shapes right of refShape, adjoin them from closest to refShape
            for (var i = refShapeIndex + 1; i < sortedShapes.Count; i++)
            {
                var neighbour = sortedShapes[i];
                var allPointsOfNeighbour = Graphics.GetRealCoordinates(neighbour);
                var topOfShape = Graphics.TopMostPoint(allPointsOfNeighbour).Y;
                neighbour.IncrementLeft(centerOfRef.X - Graphics.GetCenterPoint(neighbour).X);
                neighbour.IncrementTop(lowest - topOfShape);

                lowest = Graphics.BottomMostPoint(allPointsOfNeighbour).Y + lowest - topOfShape;
            }
        }
        #endregion

        #region Swap
        public static void Swap(List<Shape> selectedShapes)
        {
            if (selectedShapes.Count < 2)
            {
                //Error
                return;
            }

            var sortedShapes = Graphics.SortShapesByLeft(selectedShapes);
            var firstPos = Graphics.GetCenterPoint(sortedShapes[0]);

            for (var i = 0; i < sortedShapes.Count; i++)
            {
                var currentShape = sortedShapes[i];
                if (i < sortedShapes.Count - 1)
                {
                    var currentPos = Graphics.GetCenterPoint(currentShape);
                    var nextPos = Graphics.GetCenterPoint(sortedShapes[i + 1]);

                    currentShape.IncrementLeft(nextPos.X - currentPos.X);
                    currentShape.IncrementTop(nextPos.Y - currentPos.Y);
                }
                else
                {
                    var currentPos = Graphics.GetCenterPoint(currentShape);
                    currentShape.IncrementLeft(firstPos.X - currentPos.X);
                    currentShape.IncrementTop(firstPos.Y - currentPos.Y);
                }
            }
        }
        #endregion

        #region Distribute
        public static void DistributeHorizontal(List<Shape> selectedShapes, float slideWidth)
        {
            var shapeCount = selectedShapes.Count;

            var refShape = selectedShapes[0];
            var allPointsOfRef = Graphics.GetRealCoordinates(refShape);
            Drawing.PointF rightMostRef;

            if (_distributeUseSlideAsReference)
            {
                var horizontalDistanceInRef = slideWidth;
                var spaceBetweenShapes = horizontalDistanceInRef;

                foreach (var s in selectedShapes)
                {
                    var allPoints = Graphics.GetRealCoordinates(s);
                    var shapeWidth = Graphics.RealWidth(allPoints);
                    spaceBetweenShapes -= shapeWidth;
                }

                // TODO: guard against spaceBetweenShapes < 0

                spaceBetweenShapes /= shapeCount + 1;

                for (var i = 0; i < shapeCount; i++)
                {
                    var currShape = selectedShapes[i];
                    var allPoints = Graphics.GetRealCoordinates(currShape);
                    var leftMost = Graphics.LeftMostPoint(allPoints);
                    if (i == 0)
                    {
                        currShape.IncrementLeft(spaceBetweenShapes - leftMost.X);
                    }
                    else
                    {
                        refShape = selectedShapes[i - 1];
                        allPointsOfRef = Graphics.GetRealCoordinates(refShape);
                        rightMostRef = Graphics.RightMostPoint(allPointsOfRef);
                        currShape.IncrementLeft(rightMostRef.X - leftMost.X + spaceBetweenShapes);
                    }
                }
            }
            else
            {
                if (shapeCount < 2)
                {
                    //Error
                    return;
                }

                var horizontalDistanceInRef = Graphics.RealWidth(allPointsOfRef);
                var spaceBetweenShapes = horizontalDistanceInRef;

                for (var i = 1; i < shapeCount; i++)
                {
                    var s = selectedShapes[i];
                    var allPoints = Graphics.GetRealCoordinates(s);
                    var shapeWidth = Graphics.RealWidth(allPoints);
                    spaceBetweenShapes -= shapeWidth;
                }

                // TODO: guard against spaceBetweenShapes < 0

                spaceBetweenShapes /= shapeCount;

                for (var i = 1; i < shapeCount; i++)
                {
                    var currShape = selectedShapes[i];
                    var allPoints = Graphics.GetRealCoordinates(currShape);
                    var leftMost = Graphics.LeftMostPoint(allPoints);
                    refShape = selectedShapes[i - 1];
                    allPointsOfRef = Graphics.GetRealCoordinates(refShape);

                    if (i == 1)
                    {
                        var leftMostRef = Graphics.LeftMostPoint(allPointsOfRef);
                        currShape.IncrementLeft(leftMostRef.X - leftMost.X + spaceBetweenShapes);
                    }
                    else
                    {
                        rightMostRef = Graphics.RightMostPoint(allPointsOfRef);
                        currShape.IncrementLeft(rightMostRef.X - leftMost.X + spaceBetweenShapes);
                    }
                }
            }
        } 

        public static void DistributeVertical(List<Shape> selectedShapes, float slideHeight)
        {
            var shapeCount = selectedShapes.Count;

            var refShape = selectedShapes[0];
            var allPointsOfRef = Graphics.GetRealCoordinates(refShape);
            Drawing.PointF lowestRef;

            if (_distributeUseSlideAsReference)
            {
                var verticalDistanceInRef = slideHeight;
                var spaceBetweenShapes = verticalDistanceInRef;

                foreach (var s in selectedShapes)
                {
                    var allPoints = Graphics.GetRealCoordinates(s);
                    var shapeHeight = Graphics.RealHeight(allPoints);
                    spaceBetweenShapes -= shapeHeight;
                }

                // TODO: guard against spaceBetweenShapes < 0

                spaceBetweenShapes /= shapeCount + 1;

                for (var i = 0; i < shapeCount; i++)
                {
                    var currShape = selectedShapes[i];
                    var allPoints = Graphics.GetRealCoordinates(currShape);
                    var topMost = Graphics.TopMostPoint(allPoints);
                    if (i == 0)
                    {
                        currShape.IncrementTop(spaceBetweenShapes - topMost.Y);
                    }
                    else
                    {
                        refShape = selectedShapes[i - 1];
                        allPointsOfRef = Graphics.GetRealCoordinates(refShape);
                        lowestRef = Graphics.BottomMostPoint(allPointsOfRef);
                        currShape.IncrementTop(lowestRef.Y - topMost.Y + spaceBetweenShapes);
                    }
                }
            }
            else
            {
                if (shapeCount < 2)
                {
                    //Error
                    return;
                }

                var verticalDistanceInRef = Graphics.RealHeight(allPointsOfRef);
                var spaceBetweenShapes = verticalDistanceInRef;

                for (var i = 1; i < shapeCount; i++)
                {
                    var s = selectedShapes[i];
                    var allPoints = Graphics.GetRealCoordinates(s);
                    var shapeHeight = Graphics.RealHeight(allPoints);
                    spaceBetweenShapes -= shapeHeight;
                }

                // TODO: guard against spaceBetweenShapes < 0

                spaceBetweenShapes /= shapeCount;

                for (var i = 1; i < shapeCount; i++)
                {
                    var currShape = selectedShapes[i];
                    var allPoints = Graphics.GetRealCoordinates(currShape);
                    var topMost = Graphics.TopMostPoint(allPoints);
                    refShape = selectedShapes[i - 1];
                    allPointsOfRef = Graphics.GetRealCoordinates(refShape);

                    if (i == 1)
                    {
                        var topMostRef = Graphics.TopMostPoint(allPointsOfRef);
                        currShape.IncrementTop(topMostRef.Y - topMost.Y + spaceBetweenShapes);
                    }
                    else
                    {
                        lowestRef = Graphics.BottomMostPoint(allPointsOfRef);
                        currShape.IncrementTop(lowestRef.Y - topMost.Y + spaceBetweenShapes);
                    }
                }
            }
        }

        public static void DistributeCenter(List<Shape> selectedShapes, float slideWidth, float slideHeight)
        {
            DistributeHorizontal(selectedShapes, slideWidth);
            DistributeVertical(selectedShapes, slideHeight);
        }

        public static void DistributeShapes(List<Shape> selectedShapes)
        {
            var shapeCount = selectedShapes.Count;

            if (shapeCount < 2)
            {
                //Error
                return;
            }

            if (shapeCount == 2)
            {
                return;
            }

            var firstRef = selectedShapes[0];
            var lastRef = selectedShapes[selectedShapes.Count - 1];
            Shape refShape;

            var allPointsOfFirstRef = Graphics.GetRealCoordinates(firstRef);
            var allPointsOfLastRef = Graphics.GetRealCoordinates(lastRef);
            Drawing.PointF[] allPointsOfRef;

            var horizontalDistance = Graphics.LeftMostPoint(allPointsOfLastRef).X - Graphics.RightMostPoint(allPointsOfFirstRef).X;
            var verticalDistance = Graphics.TopMostPoint(allPointsOfLastRef).Y - Graphics.BottomMostPoint(allPointsOfFirstRef).Y;

            var spaceBetweenShapes = horizontalDistance;

            for (var i = 1; i < shapeCount - 1; i++)
            {
                var s = selectedShapes[i];
                var allPoints = Graphics.GetRealCoordinates(s);
                var shapeWidth = Graphics.RealWidth(allPoints);
                spaceBetweenShapes -= shapeWidth;
            }

            // TODO: guard against spaceBetweenShapes < 0

            spaceBetweenShapes /= (shapeCount-1);
            
            for (var i = 1; i < shapeCount - 1; i++)
            {
                var currShape = selectedShapes[i];
                var allPoints = Graphics.GetRealCoordinates(currShape);
                var leftMost = Graphics.LeftMostPoint(allPoints);
                refShape = selectedShapes[i - 1];
                allPointsOfRef = Graphics.GetRealCoordinates(refShape);

                var rightMostRef = Graphics.RightMostPoint(allPointsOfRef);
                currShape.IncrementLeft(rightMostRef.X - leftMost.X + spaceBetweenShapes);
            }

            spaceBetweenShapes = verticalDistance;
            for (var i = 1; i < shapeCount - 1; i++)
            {
                var s = selectedShapes[i];
                var allPoints = Graphics.GetRealCoordinates(s);
                var shapeHeight = Graphics.RealHeight(allPoints);
                spaceBetweenShapes -= shapeHeight;
            }

            // TODO: guard against spaceBetweenShapes < 0

            spaceBetweenShapes /= shapeCount;

            for (var i = 1; i < shapeCount - 1; i++)
            {
                var currShape = selectedShapes[i];
                var allPoints = Graphics.GetRealCoordinates(currShape);
                var topMost = Graphics.TopMostPoint(allPoints);
                refShape = selectedShapes[i - 1];
                allPointsOfRef = Graphics.GetRealCoordinates(refShape);

                var lowestRef = Graphics.BottomMostPoint(allPointsOfRef);
                currShape.IncrementTop(lowestRef.Y - topMost.Y + spaceBetweenShapes);
            }
        }

        public static void DistributeGrid(List<Shape> selectedShapes, int rowLength, int colLength)
        {
            var colLengthGivenFullRows = (int)Math.Ceiling((double)selectedShapes.Count / rowLength);
            if (colLength <= colLengthGivenFullRows)
            {
                DistributeGridByRow(selectedShapes, rowLength, colLength);
            }
            else
            {
                DistributeGridByCol(selectedShapes, rowLength, colLength);
            }
        }

        public static void DistributeGridByRow(List<Shape> selectedShapes, int rowLength, int colLength)
        {
            var refPoint = Graphics.GetCenterPoint(selectedShapes[0]);

            var allShapes = new List<PPShape>();
            var numShapes = selectedShapes.Count;

            for (var i = 0; i < numShapes; i++)
            {
                allShapes.Add(new PPShape(selectedShapes[i]));
            }

            var numIndicesToSkip = IndicesToSkip(numShapes, rowLength, _distributeGridAlignment);

            var rowDifferences = GetLongestWidthsOfRowsByRow(allShapes, rowLength, numIndicesToSkip);
            var colDifferences = GetLongestHeightsOfColsByRow(allShapes, rowLength, colLength);

            var posX = refPoint.X;
            var posY = refPoint.Y;
            var remainder = numShapes % rowLength;
            var differenceIndex = 0;

            for (var i = 0; i < numShapes; i++)
            {
                //Start of new row
                if (i % rowLength == 0 && i != 0)
                {
                    posX = refPoint.X;
                    differenceIndex = 0;
                    posY += GetSpaceBetweenShapes(i / rowLength - 1, i / rowLength, colDifferences, _marginTop, _marginBottom);
                }

                //If last row, offset by num of indices to skip
                if (numShapes - i == remainder)
                {
                    differenceIndex = numIndicesToSkip;
                    posX += GetSpaceBetweenShapes(0, differenceIndex, rowDifferences, _marginLeft, _marginRight);
                }

                var currentShape = selectedShapes[i];
                var center = Graphics.GetCenterPoint(currentShape);
                currentShape.IncrementLeft(posX - center.X);
                currentShape.IncrementTop(posY - center.Y);

                posX += GetSpaceBetweenShapes(differenceIndex, differenceIndex + 1, rowDifferences, _marginLeft, _marginRight);
                differenceIndex++;
            }
        }

        public static void DistributeGridByCol(List<Shape> selectedShapes, int rowLength, int colLength)
        {
            var refPoint = Graphics.GetCenterPoint(selectedShapes[0]);

            var allShapes = new List<PPShape>();
            var numShapes = selectedShapes.Count;

            for (var i = 0; i < numShapes; i++)
            {
                allShapes.Add(new PPShape(selectedShapes[i]));
            }

            var numIndicesToSkip = IndicesToSkip(numShapes, colLength, _distributeGridAlignment);

            var rowDifferences = GetLongestWidthsOfRowsByCol(allShapes, rowLength, colLength, numIndicesToSkip);
            var colDifferences = GetLongestHeightsOfColsByCol(allShapes, rowLength, colLength, numIndicesToSkip);

            var posX = refPoint.X;
            var posY = refPoint.Y;
            var remainder = colLength - (rowLength * colLength - numShapes);
            var augmentedShapeIndex = 0;

            for (var i = 0; i < numShapes; i++)
            {
                //If last index and need to skip, skip index 
                if (numIndicesToSkip > 0 && IsLastIndexOfRow(augmentedShapeIndex, rowLength))
                {
                    numIndicesToSkip--;
                    augmentedShapeIndex++;
                }

                //If last index and no more remainder, skip the rest
                if (IsLastIndexOfRow(augmentedShapeIndex, rowLength))
                {
                    if (remainder <= 0)
                    {
                        augmentedShapeIndex++;
                    }
                    else
                    {
                        remainder--;
                    }
                }

                if (IsFirstIndexOfRow(augmentedShapeIndex, rowLength) && augmentedShapeIndex != 0)
                {
                    posX = refPoint.X;
                    posY += GetSpaceBetweenShapes(augmentedShapeIndex / rowLength - 1, augmentedShapeIndex / rowLength, colDifferences, _marginTop, _marginBottom);
                }

                var currentShape = selectedShapes[i];
                var center = Graphics.GetCenterPoint(currentShape);
                currentShape.IncrementLeft(posX - center.X);
                currentShape.IncrementTop(posY - center.Y);

                posX += GetSpaceBetweenShapes(augmentedShapeIndex % rowLength, augmentedShapeIndex % rowLength + 1, rowDifferences, _marginLeft, _marginRight);
                augmentedShapeIndex++;
            }
        }
        #endregion

        #endregion

        #region Util
        public static double AngleBetweenTwoPoints(Drawing.PointF refPoint, Drawing.PointF pt)
        {
            var angle = Math.Atan((pt.Y - refPoint.Y) / (pt.X - refPoint.X)) * 180 / Math.PI;

            if (pt.X - refPoint.X > 0)
            {
                angle = 90 + angle;
            }
            else
            {
                angle = 270 + angle;
            }

            return angle;
        }

        public static bool NearlyEqual(float a, float b, float epsilon)
        {
            var absA = Math.Abs(a);
            var absB = Math.Abs(b);
            var diff = Math.Abs(a - b);

            if (a == b)
            { // shortcut, handles infinities
                return true;
            }
            if (a == 0 || b == 0 || diff < float.Epsilon)
            {
                // a or b is zero or both are extremely close to it
                // relative error is less meaningful here
                return diff < epsilon;
            }
            // use relative error
            return diff / (absA + absB) < epsilon;
        }

        private static int GetDirectionWrtRefShape(Shape shape, float angleFromRefShape)
        {
            float defaultUpAngle;
            var hasDefaultDirection = shapeDefaultUpAngle.TryGetValue(shape.AutoShapeType, out defaultUpAngle);

            if (shape.AutoShapeType == AutoShape.msoShapeLightningBolt)
            {
                Debug.WriteLine("defaultDir: " + hasDefaultDirection);
                Debug.WriteLine("defaultAngle: " + defaultUpAngle);
            }

            if (!hasDefaultDirection)
            {
                if (IsVertical(shape))
                {
                    defaultUpAngle = 0;
                }
                else
                {
                    defaultUpAngle = 90;
                }
            }

            var angle = AddAngles(angleFromRefShape, defaultUpAngle);
            var diff = SubtractAngles(shape.Rotation, angle);
            var phaseInFloat = diff / 90;

            if (shape.AutoShapeType == AutoShape.msoShapeLightningBolt)
            {
                Debug.WriteLine("angle: " + angle);
                Debug.WriteLine("diff: " + diff);
                Debug.WriteLine("phaseInFloat: " + defaultUpAngle);
                Debug.WriteLine("equal: " + NearlyEqual(phaseInFloat, (float)Math.Round(phaseInFloat), Epsilon));
            }

            if (!NearlyEqual(phaseInFloat, (float)Math.Round(phaseInFloat), Epsilon))
            {
                return None;
            }

            var phase = (int)Math.Round(phaseInFloat);

            if (!hasDefaultDirection)
            {
                if (phase == Left || phase == Right)
                {
                    return Leftorright;
                }

                return Upordown;
            }

            return phase;
        }

        private static bool IsSameDirection(int a, int b)
        {
            if (a == b) return true;
            if (a == Leftorright) return b == Left || b == Right;
            if (b == Leftorright) return a == Left || a == Right;
            if (a == Upordown) return b == Up || b == Down;
            if (b == Upordown) return a == Up || a == Down;

            return false;
       }

        public static float AddAngles(float a, float b)
        {
            return (a + b) % 360;
        }

        public static float SubtractAngles(float a, float b)
        {
            var diff = a - b;
            if (diff < 0)
            {
                return 360 + diff;
            }

            return diff;
        }

        public static float[] GetLongestWidthsOfRowsByRow(List<PPShape> shapes, int rowLength, int numIndicesToSkip)
        {
            var longestWidths = new float[rowLength];
            var numShapes = shapes.Count;
            var remainder = numShapes % rowLength;

            for (var i = 0; i < numShapes; i++)
            {
                var longestRowIndex = i % rowLength;
                if (numShapes - i == remainder - 1)
                {
                    longestRowIndex += numIndicesToSkip;
                }
                if (longestWidths[longestRowIndex] < shapes[i].AbsoluteWidth)
                {
                    longestWidths[longestRowIndex] = shapes[i].AbsoluteWidth;
                }
            }

            return longestWidths;
        }

        public static float[] GetLongestHeightsOfColsByRow(List<PPShape> shapes, int rowLength, int colLength)
        {
            var longestHeights = new float[colLength];

            for (var i = 0; i < shapes.Count; i++)
            {
                var longestHeightIndex = i / rowLength;
                if (longestHeights[longestHeightIndex] < shapes[i].AbsoluteHeight)
                {
                    longestHeights[longestHeightIndex] = shapes[i].AbsoluteHeight;
                }
            }

            return longestHeights;
        }

        public static float[] GetLongestWidthsOfRowsByCol(List<PPShape> shapes, int rowLength, int colLength, int numIndicesToSkip)
        {
            var longestWidths = new float[rowLength];
            var numShapes = shapes.Count;
            var augmentedShapeIndex = 0;
            var remainder = colLength - (rowLength * colLength - numShapes);

            for (var i = 0; i < numShapes; i++)
            {
                //If last index and need to skip, skip index 
                if (numIndicesToSkip > 0 && IsLastIndexOfRow(augmentedShapeIndex, rowLength))
                {
                    numIndicesToSkip--;
                    augmentedShapeIndex++;
                }

                //If last index and no more remainder, skip the rest
                if (IsLastIndexOfRow(augmentedShapeIndex, rowLength))
                {
                    if (remainder <= 0)
                    {
                        augmentedShapeIndex++;
                    }
                    else
                    {
                        remainder--;
                    }
                }

                var longestWidthsArrayIndex = augmentedShapeIndex % rowLength;

                if (longestWidths[longestWidthsArrayIndex] < shapes[i].AbsoluteWidth)
                {
                    longestWidths[longestWidthsArrayIndex] = shapes[i].AbsoluteWidth;
                }

                augmentedShapeIndex++;
            }

            return longestWidths;
        }

        public static float[] GetLongestHeightsOfColsByCol(List<PPShape> shapes, int rowLength, int colLength, int numIndicesToSkip)
        {
            var longestHeights = new float[colLength];
            var numShapes = shapes.Count;
            var augmentedShapeIndex = 0;
            var remainder = colLength - (rowLength * colLength - numShapes);

            for (var i = 0; i < numShapes; i++)
            {              
                //If last index and need to skip, skip index 
                if (numIndicesToSkip > 0 && IsLastIndexOfRow(augmentedShapeIndex, rowLength))
                {
                    numIndicesToSkip--;
                    augmentedShapeIndex++;
                }

                //If last index and no more remainder, skip the rest
                if (IsLastIndexOfRow(augmentedShapeIndex, rowLength))
                {
                    if (remainder <= 0)
                    {
                        augmentedShapeIndex++;
                    }
                    else
                    {
                        remainder--;
                    }
                }

                var longestHeightArrayIndex = augmentedShapeIndex / rowLength;

                if (longestHeights[longestHeightArrayIndex] < shapes[i].AbsoluteHeight)
                {
                    longestHeights[longestHeightArrayIndex] = shapes[i].AbsoluteHeight;
                }

                augmentedShapeIndex++;
            }

            return longestHeights;
        }

        private static bool IsFirstIndexOfRow(int index, int rowLength)
        {
            return index % rowLength == 0;
        }

        private static bool IsLastIndexOfRow(int index, int rowLength)
        {
            return index % rowLength == rowLength - 1;
        }

        public static int IndicesToSkip(int totalSelectedShapes, int rowLength, GridAlignment alignment)
        {
            var numOfShapesInLastRow = totalSelectedShapes % rowLength;

            if (alignment == GridAlignment.AlignLeft || numOfShapesInLastRow == 0)
            {
                return 0;
            }

            if (alignment == GridAlignment.AlignRight)
            {
                return rowLength - numOfShapesInLastRow;
            }

            if (alignment == GridAlignment.AlignCenter)
            {
                var difference = rowLength - numOfShapesInLastRow;
                return difference / 2;
            }

            return 0;
        }

        private static float GetSpaceBetweenShapes(int index1, int index2, float[] differences, float margin1, float margin2)
        {
            if (index1 >= differences.Length || index2 >= differences.Length)
            {
                return -1;
            }

            var start = 0;
            var end = 0;

            if (index1 < index2)
            {
                start = index1;
                end = index2;
            }
            else
            {
                start = index2;
                end = index1;
            }

            float difference = 0;

            for (var i = index1; i < index2; i++)
            {
                difference += (differences[i] / 2 + margin1 + margin2 + differences[i + 1] / 2);
            }

            return difference;
        }

        private static void Init()
        {
            shapeDefaultUpAngle = new Dictionary<MsoAutoShapeType, float>();

            shapeDefaultUpAngle.Add(AutoShape.msoShapeLeftArrow, RotateLeft);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeLeftRightArrow, RotateLeft);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeLeftArrowCallout, RotateLeft);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeLeftRightArrowCallout, RotateLeft);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeCurvedLeftArrow, RotateLeft);

            shapeDefaultUpAngle.Add(AutoShape.msoShapeRightArrow, RotateRight);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeBentArrow, RotateRight);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeStripedRightArrow, RotateRight);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeNotchedRightArrow, RotateRight);
            shapeDefaultUpAngle.Add(AutoShape.msoShapePentagon, RotateRight);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeChevron, RotateRight);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeRightArrowCallout, RotateRight);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeCurvedRightArrow, RotateRight);

            shapeDefaultUpAngle.Add(AutoShape.msoShapeUpArrow, RotateUp);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeBentUpArrow, RotateUp);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeUpDownArrow, RotateUp);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeLeftRightUpArrow, RotateUp);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeLeftUpArrow, RotateUp);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeUpArrowCallout, RotateUp);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeCurvedUpArrow, RotateUp);

            shapeDefaultUpAngle.Add(AutoShape.msoShapeDownArrow, RotateDown);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeUTurnArrow, RotateDown);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeDownArrowCallout, RotateDown);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeCurvedDownArrow, RotateDown);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeCircularArrow, RotateDown);
        }

        #endregion

        /// <summary>
        ///     Method returns the internal (as opposed to bounding box) vertices of the shape.
        /// </summary>
        /// <param name="selectedShape">Shape to get the internal vertices</param>
        /// <param name="currentSlide">Slide that the shape exist in</param>
        /// <returns> 
        ///     An array of 4 float values. They are arranged as such from left to right in the array: 
        ///     left most, top most, right most, bottom most.
        ///     In the event that the method fails to get the internal vertices of the shape, 
        ///     it will return the vertices of the bounding box instead.
        /// </returns>
        public static float[] GetInternalVertices(Shape selectedShape, PowerPointSlide currentSlide)
        {
            var result = new float[4];

            // Get the temp folder path for temporary storage of image
            var folderPath = Path.GetTempPath() + @"\" + "pptlabs_positionsLab" + @"\";
            // Create temp folder if it does not exist
            if (!Directory.Exists(folderPath))
            {
                try
                {
                    Directory.CreateDirectory(folderPath);
                }
                catch (Exception)
                {
                    // TODO: inform user of failure to create directory
                    // TODO: Return bounding box vertices
                    return result;
                }
            }
            var currShape = selectedShape;
            
            // Checks to ensure that it's possible to get internal vertices
            if (ShapeFillsBoundingBox.Contains(currShape.AutoShapeType) ||
                ShapeHasNoNodes.Contains(currShape.AutoShapeType) ||
                ShapeNotSupportedForInternalVertices.Contains(currShape.AutoShapeType)||
                currShape.Connector == MsoTriState.msoTrue)
            {
                // TODO: Return bounding box vertices
                return result;
            }

            // Save the rotation of the shape. The shape's rotation is then set to 0 to facilitate
            // importing the image into the exact position of the shape. The rotation is then
            // applied to the imported image after importing
            var rotationAngle = currShape.Rotation;
            currShape.Rotation = 0;

            currShape.Copy();           // DEBUG

            // Export the image into the temp folder, and get the path where it is stored
            var imagePath = ExportShape(folderPath, currShape, PpShapeFormat.ppShapeFormatEMF);

            // Import the image into current slide
            var importedImage = ImportImage(imagePath, currentSlide, currShape);

            // Ungroup image to get freeform shapes within
            var ungroupedShapes = importedImage.Ungroup();
            
            // Apply rotation to the image and original shape again
            ungroupedShapes.Rotation = rotationAngle;
            currShape.Rotation = rotationAngle;

            var shapesInImage = ungroupedShapes.GroupItems;

            // The freeform shape that's the boundary is always the third shape
            var nodes = shapesInImage[3].Nodes;
            Debug.WriteLine("Processing shape: " + currShape.Name + "'s nodes");        //DEBUG


            var leftMost = new Drawing.PointF();
            var topMost = new Drawing.PointF();
            var rightMost = new Drawing.PointF();
            var bottomMost = new Drawing.PointF();

            // Initialising topMost and leftMost to the first node
            leftMost.X = nodes[1].Points[1, 1];
            leftMost.Y = nodes[1].Points[1, 2];
            Debug.WriteLine("Initial left most: (" + leftMost.X + ", " + leftMost.Y + ")");     //DEBUG
            topMost.X = nodes[1].Points[1, 1];
            topMost.Y = nodes[1].Points[1, 2];
            Debug.WriteLine("Initial top most: (" + topMost.X + ", " + topMost.Y + ")");        //DEBUG

            // Get the co-ordinates of node in that freeform shape
            try
            {
                foreach (PowerPoint.ShapeNode sn in nodes)
                {
                    // Works for most shapes, except for rare cases like Chords 
                    if (sn.EditingType == MsoEditingType.msoEditingSymmetric)
                    {
                        continue;
                    }

                    float x = sn.Points[1, 1];
                    float y = sn.Points[1, 2];

                    Debug.WriteLine("Co-ord : (" + x + " ," + y + ")");             //DEBUG

                    if (x < leftMost.X)
                    {
                        leftMost.X = x;
                        leftMost.Y = y;
                    }
                    if (y < topMost.Y)
                    {
                        topMost.X = x;
                        topMost.Y = y;
                    }
                    if (x > rightMost.X)
                    {
                        rightMost.X = x;
                        rightMost.Y = y;
                    }
                    if (y > bottomMost.Y)
                    {
                        bottomMost.X = x;
                        bottomMost.Y = y;
                    }
                }
            }
            catch (Exception)
            {
                Debug.WriteLine("Failed to process nodes for: " + currShape.Name);
            }
            
            // DEBUG
            Debug.WriteLine("Left Most: (" + leftMost.X + ", " + leftMost.Y + ")");
            Debug.WriteLine("Top Most: (" + topMost.X + ", " + topMost.Y + ")");
            Debug.WriteLine("Right Most: (" + rightMost.X + ", " + rightMost.Y + ")");
            Debug.WriteLine("Bottom Most: (" + bottomMost.X + ", " + bottomMost.Y + ")");

            // DEBUG Pasting the original shape to manually check where are the extreme ends
            var pic = currentSlide.Shapes.PasteSpecial(PpPasteDataType.ppPastePNG)[1];
            pic.Left = leftMost.X;
            pic.Top = leftMost.Y;
            pic.Name = "leftMost" + Guid.NewGuid();

            pic = currentSlide.Shapes.PasteSpecial(PpPasteDataType.ppPastePNG)[1];
            pic.Left = topMost.X;
            pic.Top = topMost.Y;
            pic.Name = "topMost " + Guid.NewGuid();

            pic = currentSlide.Shapes.PasteSpecial(PpPasteDataType.ppPastePNG)[1];
            pic.Left = rightMost.X;
            pic.Top = rightMost.Y;
            pic.Name = "rightMost" + Guid.NewGuid();

            pic = currentSlide.Shapes.PasteSpecial(PpPasteDataType.ppPastePNG)[1];
            pic.Left = bottomMost.X;
            pic.Top = bottomMost.Y;
            pic.Name = "bottomMost" + Guid.NewGuid();

            result[0] = leftMost.X;
            result[1] = topMost.Y;
            result[2] = rightMost.X;
            result[3] = bottomMost.Y;

            // Clean the temp directory
            CleanTempDirectory(folderPath);
            return result;
        }

        /// <summary>
        /// Exports a set of shapes as one image with the specified format
        /// </summary>
        /// <param name="folderPath">The path where the image is to be exported to</param>
        /// <param name="selectedShape">The selected shapes to export</param>
        /// <param name="format">The format of image to be exported as. Should be a PpShapeFormat enum value</param>
        private static string ExportShape(string folderPath, Shape selectedShape, PpShapeFormat format)
        {
            var imagePath = folderPath + "exportedImage" + DateTime.Now.GetHashCode() + "-" + Guid.NewGuid().ToString().Substring(0, 7);
            selectedShape.Export(imagePath + ".emf", format, 1, 1);

            return imagePath + ".emf";
        }

        /// <summary>
        /// Imports an image from the specified path
        /// </summary>
        /// <param name="imagePath">Path where the image to be imported can be found</param>
        /// <param name="slideToImportTo"></param>
        /// <param name="originalShape"></param>
        /// <returns></returns>
        private static Shape ImportImage(string imagePath, PowerPointSlide slideToImportTo, Shape originalShape)
        {
            slideToImportTo.Shapes.AddPicture(imagePath, MsoTriState.msoFalse, MsoTriState.msoTrue, 0, 0);
            var importedShape = slideToImportTo.Shapes[slideToImportTo.Shapes.Count];

            Debug.WriteLine("Imported shape holding on to shape: " + importedShape.Name);
            var originalCenter = Graphics.GetCenterPoint(originalShape);
            var importedCenter = Graphics.GetCenterPoint(importedShape);

            importedShape.IncrementLeft(originalCenter.X - importedCenter.X);
            importedShape.IncrementTop(originalCenter.Y - importedCenter.Y);

            return importedShape;
        }

        /// <summary>
        /// Cleans the temporary directory used to store images
        /// </summary>
        /// <param name="folderPath">Path for the temp folder</param>
        private static void CleanTempDirectory(string folderPath)
        {
            var directory = new DirectoryInfo(folderPath);

            try
            {
                foreach (var file in directory.GetFiles()) file.Delete();
                foreach (var subDirectory in directory.GetDirectories()) subDirectory.Delete(true);
            }
            catch (Exception)
            {
                Debug.WriteLine("Unable to delete");
                // ignore ex, if cannot delete trash
            }
        }
    }
}
