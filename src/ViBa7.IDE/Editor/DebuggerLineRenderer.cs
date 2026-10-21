using System;
using System.IO;
using System.Windows;
using System.Windows.Media;
using ICSharpCode.AvalonEdit;
using ICSharpCode.AvalonEdit.Rendering;
using ViBa7.ProjectModel.Navigation;

namespace ViBa7.IDE.Editor;

/// <summary>
/// Renderizador de fondo para líneas de depuración en AvalonEdit.
/// Resalta sutilmente las líneas con puntos de interrupción activos (rojo transparente)
/// y destaca en amarillo la línea de instrucción en pausa activa durante la depuración interactiva.
/// </summary>
public class DebuggerLineRenderer : IBackgroundRenderer
{
    private readonly TextEditor _editor;
    private readonly string _filePath;

    public KnownLayer Layer => KnownLayer.Background;

    // Colores para punto de interrupción
    private static readonly Brush BreakpointLineBrush = new SolidColorBrush(Color.FromArgb(32, 208, 32, 32));
    private static readonly Pen BreakpointLinePen = new Pen(new SolidColorBrush(Color.FromArgb(64, 180, 20, 20)), 1);

    // Colores para instrucción activa (Pausa / Step / Breakpoint Hit)
    private static readonly Brush PausedLineBrush = new SolidColorBrush(Color.FromArgb(90, 255, 234, 0));
    private static readonly Pen PausedLinePen = new Pen(new SolidColorBrush(Color.FromRgb(218, 175, 0)), 1.2);

    static DebuggerLineRenderer()
    {
        BreakpointLineBrush.Freeze();
        BreakpointLinePen.Freeze();
        PausedLineBrush.Freeze();
        PausedLinePen.Freeze();
    }

    public DebuggerLineRenderer(TextEditor editor, string filePath)
    {
        _editor = editor ?? throw new ArgumentNullException(nameof(editor));
        _filePath = filePath ?? string.Empty;

        BreakpointManager.Instance.BreakpointsChanged += OnDebuggerVisualsChanged;
        BookmarkMargin.ExecutionPointerChanged += OnDebuggerVisualsChanged;
    }

    private void OnDebuggerVisualsChanged()
    {
        _editor.Dispatcher.InvokeAsync(() =>
        {
            _editor.TextArea?.TextView?.InvalidateLayer(KnownLayer.Background);
        });
    }

    public void Draw(TextView textView, DrawingContext drawingContext)
    {
        if (_editor.Document == null || _editor.TextArea == null || !textView.VisualLinesValid)
            return;

        string? execFile = BookmarkMargin.ActiveExecutionFilePath;
        int execLine = BookmarkMargin.ActiveExecutionLineNumber;
        bool isExecutionFile = !string.IsNullOrEmpty(execFile) &&
                               (string.Equals(_filePath, execFile, StringComparison.OrdinalIgnoreCase) ||
                                string.Equals(Path.GetFileName(_filePath), Path.GetFileName(execFile), StringComparison.OrdinalIgnoreCase));

        double width = Math.Max(textView.ActualWidth, textView.RenderSize.Width);
        if (width <= 0) width = 2000;

        foreach (var visualLine in textView.VisualLines)
        {
            int lineNumber = visualLine.FirstDocumentLine.LineNumber;
            double y = visualLine.VisualTop - textView.VerticalOffset;
            double height = visualLine.Height;
            var rect = new Rect(0, y, width, height);

            bool isExecutionLine = isExecutionFile && lineNumber == execLine;
            bool isBreakpointLine = BreakpointManager.Instance.HasBreakpoint(_filePath, lineNumber);

            // 1. Línea activa de ejecución (Amarillo - máxima prioridad visual)
            if (isExecutionLine)
            {
                drawingContext.DrawRectangle(PausedLineBrush, PausedLinePen, rect);
            }
            // 2. Línea con punto de interrupción (Rojo sutil)
            else if (isBreakpointLine)
            {
                drawingContext.DrawRectangle(BreakpointLineBrush, BreakpointLinePen, rect);
            }
        }
    }
}
