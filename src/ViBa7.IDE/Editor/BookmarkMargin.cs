using System;
using System.IO;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using ICSharpCode.AvalonEdit;
using ICSharpCode.AvalonEdit.Editing;
using ICSharpCode.AvalonEdit.Rendering;
using ViBa7.ProjectModel.Navigation;

namespace ViBa7.IDE.Editor;

/// <summary>
/// Margen interactivo de puntos de interrupción, ejecución y marcadores (Glyph Margin / Gutter)
/// ubicado en el extremo izquierdo absoluto del editor de código.
/// Dibuja círculos rojos de Breakpoint, cintas de Bookmark y la flecha amarilla de instrucción activa en depuración.
/// </summary>
public class BookmarkMargin : AbstractMargin
{
    private readonly TextEditor _editor;
    private readonly string _filePath;

    private static readonly Brush GutterBackground = new SolidColorBrush(Color.FromRgb(246, 246, 246));
    private static readonly Pen GutterSeparator = new Pen(new SolidColorBrush(Color.FromRgb(226, 226, 226)), 1);

    // Marcadores (Bookmarks - Azul)
    private static readonly Brush BookmarkBrush = new SolidColorBrush(Color.FromRgb(0, 122, 204));
    private static readonly Pen BookmarkPen = new Pen(new SolidColorBrush(Color.FromRgb(0, 90, 160)), 1);

    // Puntos de interrupción (Breakpoints - Rojo carmesí clásico)
    private static readonly Brush BreakpointBrush = new SolidColorBrush(Color.FromRgb(208, 32, 32));
    private static readonly Pen BreakpointPen = new Pen(new SolidColorBrush(Color.FromRgb(140, 16, 16)), 1.2);

    // Flecha de ejecución activa (Amarillo dorado / Visual Studio pointer)
    private static readonly Brush ExecutionArrowBrush = new SolidColorBrush(Color.FromRgb(255, 215, 0));
    private static readonly Pen ExecutionArrowPen = new Pen(new SolidColorBrush(Color.FromRgb(154, 125, 10)), 1);

    // Estado global de la instrucción en pausa del depurador
    private static string? _activeExecutionFilePath;
    private static int _activeExecutionLineNumber;
    public static event Action? ExecutionPointerChanged;

    public static void SetExecutionPointer(string? filePath, int lineNumber)
    {
        _activeExecutionFilePath = filePath;
        _activeExecutionLineNumber = lineNumber;
        ExecutionPointerChanged?.Invoke();
    }

    public static void ClearExecutionPointer()
    {
        _activeExecutionFilePath = null;
        _activeExecutionLineNumber = 0;
        ExecutionPointerChanged?.Invoke();
    }

    public static string? ActiveExecutionFilePath => _activeExecutionFilePath;
    public static int ActiveExecutionLineNumber => _activeExecutionLineNumber;

    static BookmarkMargin()
    {
        GutterBackground.Freeze();
        GutterSeparator.Freeze();
        BookmarkBrush.Freeze();
        BookmarkPen.Freeze();
        BreakpointBrush.Freeze();
        BreakpointPen.Freeze();
        ExecutionArrowBrush.Freeze();
        ExecutionArrowPen.Freeze();
    }

    public BookmarkMargin(TextEditor editor, string filePath)
    {
        _editor = editor ?? throw new ArgumentNullException(nameof(editor));
        _filePath = filePath ?? string.Empty;

        Cursor = Cursors.Hand;
        ToolTip = "Clic izquierdo para alternar Punto de Interrupción (F9).";

        BookmarkManager.Instance.BookmarksChanged += OnGutterChanged;
        BreakpointManager.Instance.BreakpointsChanged += OnGutterChanged;
        ExecutionPointerChanged += OnGutterChanged;
    }

    private void OnGutterChanged()
    {
        Dispatcher.InvokeAsync(InvalidateVisual);
    }

    protected override void OnTextViewChanged(TextView oldTextView, TextView newTextView)
    {
        if (oldTextView != null)
        {
            oldTextView.VisualLinesChanged -= TextView_Updated;
            oldTextView.ScrollOffsetChanged -= TextView_Updated;
        }
        base.OnTextViewChanged(oldTextView, newTextView);
        if (newTextView != null)
        {
            newTextView.VisualLinesChanged += TextView_Updated;
            newTextView.ScrollOffsetChanged += TextView_Updated;
        }
        InvalidateVisual();
    }

    private void TextView_Updated(object? sender, EventArgs e)
    {
        InvalidateVisual();
    }

    protected override Size MeasureOverride(Size availableSize)
    {
        return new Size(22, 0);
    }

    protected override void OnRender(DrawingContext drawingContext)
    {
        // 1. Fondo sutil del margen (Gutter)
        drawingContext.DrawRectangle(
            GutterBackground,
            null,
            new Rect(0, 0, RenderSize.Width, RenderSize.Height));

        // 2. Línea sutil separadora derecha
        drawingContext.DrawLine(
            GutterSeparator,
            new Point(RenderSize.Width - 0.5, 0),
            new Point(RenderSize.Width - 0.5, RenderSize.Height));

        var textView = TextView;
        if (textView == null || !textView.VisualLinesValid)
            return;

        bool hasAnyBookmarks = BookmarkManager.Instance.GetBookmarks(_filePath).Count > 0;
        bool hasAnyBreakpoints = BreakpointManager.Instance.GetBreakpoints(_filePath).Count > 0;
        bool isExecutionFile = !string.IsNullOrEmpty(_activeExecutionFilePath) &&
                               (string.Equals(_filePath, _activeExecutionFilePath, StringComparison.OrdinalIgnoreCase) ||
                                string.Equals(Path.GetFileName(_filePath), Path.GetFileName(_activeExecutionFilePath), StringComparison.OrdinalIgnoreCase));

        if (!hasAnyBookmarks && !hasAnyBreakpoints && !isExecutionFile)
            return;

        // 3. Dibujar glifos alineados con cada línea visual
        foreach (var visualLine in textView.VisualLines)
        {
            int lineNumber = visualLine.FirstDocumentLine.LineNumber;
            double y = visualLine.VisualTop - textView.VerticalOffset;
            double size = 12;
            double x = Math.Round((RenderSize.Width - size) / 2.0);
            double yOffset = Math.Round(y + (visualLine.Height - size) / 2.0);

            bool isBreakpoint = BreakpointManager.Instance.HasBreakpoint(_filePath, lineNumber);
            bool isBookmark = BookmarkManager.Instance.HasBookmark(_filePath, lineNumber);
            bool isExecutionLine = isExecutionFile && lineNumber == _activeExecutionLineNumber;

            // A. Punto de Interrupción (Círculo rojo)
            if (isBreakpoint)
            {
                drawingContext.DrawEllipse(
                    BreakpointBrush,
                    BreakpointPen,
                    new Point(x + size / 2.0, yOffset + size / 2.0),
                    size / 2.0,
                    size / 2.0);
            }

            // B. Marcador (Cinta azul en esquina inferior derecha si coexiste con breakpoint, o central)
            if (isBookmark)
            {
                if (!isBreakpoint)
                {
                    var pathGeom = new PathGeometry();
                    var figure = new PathFigure
                    {
                        StartPoint = new Point(x, yOffset),
                        IsClosed = true,
                        IsFilled = true
                    };
                    figure.Segments.Add(new LineSegment(new Point(x + size, yOffset), false));
                    figure.Segments.Add(new LineSegment(new Point(x + size, yOffset + size), false));
                    figure.Segments.Add(new LineSegment(new Point(x + size / 2.0, yOffset + size - 3), false));
                    figure.Segments.Add(new LineSegment(new Point(x, yOffset + size), false));
                    pathGeom.Figures.Add(figure);

                    drawingContext.DrawGeometry(BookmarkBrush, BookmarkPen, pathGeom);
                }
                else
                {
                    // Mini insignia azul superpuesta
                    drawingContext.DrawRectangle(BookmarkBrush, BookmarkPen, new Rect(x + size - 4, yOffset, 4, 6));
                }
            }

            // C. Flecha de ejecución activa (Amarillo dorado con punta a la derecha)
            if (isExecutionLine)
            {
                var arrowGeom = new PathGeometry();
                var arrowFig = new PathFigure
                {
                    StartPoint = new Point(x - 2, yOffset + 2),
                    IsClosed = true,
                    IsFilled = true
                };
                arrowFig.Segments.Add(new LineSegment(new Point(x + 5, yOffset + 2), false));
                arrowFig.Segments.Add(new LineSegment(new Point(x + 5, yOffset), false));
                arrowFig.Segments.Add(new LineSegment(new Point(x + size + 3, yOffset + size / 2.0), false));
                arrowFig.Segments.Add(new LineSegment(new Point(x + 5, yOffset + size), false));
                arrowFig.Segments.Add(new LineSegment(new Point(x + 5, yOffset + size - 2), false));
                arrowFig.Segments.Add(new LineSegment(new Point(x - 2, yOffset + size - 2), false));
                arrowGeom.Figures.Add(arrowFig);

                drawingContext.DrawGeometry(ExecutionArrowBrush, ExecutionArrowPen, arrowGeom);
            }
        }
    }

    protected override void OnMouseDown(MouseButtonEventArgs e)
    {
        base.OnMouseDown(e);

        if (e.ChangedButton == MouseButton.Left && TextView != null)
        {
            var pos = e.GetPosition(this);
            var visualLine = TextView.GetVisualLineFromVisualTop(pos.Y + TextView.VerticalOffset);
            if (visualLine != null)
            {
                int lineNumber = visualLine.FirstDocumentLine.LineNumber;
                // Clic izquierdo en el margen alterna punto de interrupción (estándar VB6 / Visual Studio)
                BreakpointManager.Instance.ToggleBreakpoint(_filePath, lineNumber);
                e.Handled = true;
                InvalidateVisual();
            }
        }
    }
}
