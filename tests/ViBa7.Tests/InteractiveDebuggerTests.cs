using System;
using System.Collections.Generic;
using ViBa7.ProjectModel.Execution;
using ViBa7.ProjectModel.Navigation;
using ViBa7.ProjectModel.Vbp;
using Xunit;

namespace ViBa7.Tests;

public class InteractiveDebuggerTests
{
    [Fact]
    public void BreakpointManager_TogglesAndClearsBreakpointsSuccessfully()
    {
        var manager = BreakpointManager.Instance;
        manager.ClearAll();

        string file = @"C:\TestProject\Form1.frm";

        Assert.False(manager.HasBreakpoint(file, 10));
        Assert.Equal(0, manager.TotalBreakpointCount);

        // Toggle ON
        bool added = manager.ToggleBreakpoint(file, 10);
        Assert.True(added);
        Assert.True(manager.HasBreakpoint(file, 10));
        Assert.Equal(1, manager.TotalBreakpointCount);

        // Toggle otro
        manager.ToggleBreakpoint(file, 25);
        Assert.Equal(2, manager.TotalBreakpointCount);

        var list = manager.GetBreakpoints(file);
        Assert.Contains(10, list);
        Assert.Contains(25, list);

        // Toggle OFF el primero
        bool removed = !manager.ToggleBreakpoint(file, 10);
        Assert.True(removed);
        Assert.False(manager.HasBreakpoint(file, 10));
        Assert.Equal(1, manager.TotalBreakpointCount);

        // Limpiar todos
        manager.ClearAll();
        Assert.Equal(0, manager.TotalBreakpointCount);
    }

    [Fact]
    public void Interpreter_PausesOnActiveBreakpoint_AndInspectsLocalsAndCallStack()
    {
        var ctx = new VbExecutionContext();
        ctx.ModuleVariables["g_result"] = VbValue.FromLong(0);
        var host = new LiveExecutionHost(ctx);

        string file = @"C:\App\Module1.bas";
        string code = @"
Sub Main()
    Dim x As Integer
    Dim y As Integer
    x = 10
    y = 20
    g_result = x + y
End Sub
";
        host.LoadSourceCode(code, "Module1", file);

        // Colocar breakpoint en la línea 6 ("y = 20")
        int targetLine = 6;
        ctx.Breakpoints.Add((file, targetLine));

        bool pauseHit = false;
        string? pausedAtFile = null;
        int pausedAtLine = 0;
        VbValue localXAtPause = VbValue.Empty;
        int stackDepthAtPause = 0;

        ctx.PauseHook = () =>
        {
            pauseHit = true;
            pausedAtFile = ctx.PausedFilePath;
            pausedAtLine = ctx.PausedLineNumber;
            localXAtPause = ctx.GetVariable("x");
            stackDepthAtPause = ctx.CallStack.Count;

            // Retorna true para continuar la ejecución tras la inspección
            return true;
        };

        var sig = host.InvokeProcedure("Main", "Module1");

        Assert.True(pauseHit, "El depurador debió pausar la ejecución en el punto de interrupción");
        Assert.Equal(file, pausedAtFile);
        Assert.Equal(targetLine, pausedAtLine);
        Assert.Equal(1, stackDepthAtPause);
        Assert.Equal(10, localXAtPause.AsLong()); // x ya fue asignado (x = 10) antes de y = 20
        Assert.Equal(FlowSignal.Normal, sig);
        Assert.Equal(30, ctx.GetVariable("g_result").AsLong()); // Al terminar, g_result = 10 + 20 = 30
    }

    [Fact]
    public void Interpreter_StepInto_ExecutesInstructionByInstruction()
    {
        var ctx = new VbExecutionContext();
        var host = new LiveExecutionHost(ctx);

        string file = @"C:\App\Module1.bas";
        string code = @"
Sub Pasos()
    Dim a As Integer
    a = 1
    a = 2
    a = 3
End Sub
";
        host.LoadSourceCode(code, "Module1", file);

        var linesExecuted = new List<int>();

        ctx.CurrentStepMode = StepMode.StepInto;
        ctx.PauseHook = () =>
        {
            linesExecuted.Add(ctx.PausedLineNumber);
            // Mantener StepInto activo para el siguiente paso
            ctx.CurrentStepMode = StepMode.StepInto;
            return true;
        };

        host.InvokeProcedure("Pasos", "Module1");

        // Debe haberse detenido en las 3 asignaciones ejecutables del procedimiento
        Assert.True(linesExecuted.Count >= 3, $"Se esperaban al menos 3 pausas de StepInto, se obtuvieron {linesExecuted.Count}");
    }

    [Fact]
    public void Interpreter_Abort_StopsExecutionImmediately()
    {
        var ctx = new VbExecutionContext();
        var host = new LiveExecutionHost(ctx);

        string file = @"C:\App\Module1.bas";
        string code = @"
Sub BucleInfinito()
    Dim i As Long
    i = 0
    Do While True
        i = i + 1
    Loop
End Sub
";
        host.LoadSourceCode(code, "Module1", file);

        int iterations = 0;
        ctx.CurrentStepMode = StepMode.StepInto;
        ctx.PauseHook = () =>
        {
            iterations++;
            if (iterations >= 3)
            {
                // Usuario presiona "Detener" (Shift+F5)
                ctx.State = ExecutionState.Stopped;
                return false; // Abortar
            }
            return true;
        };

        var sig = host.InvokeProcedure("BucleInfinito", "Module1");

        Assert.Equal(FlowSignal.ExitSub, sig);
        Assert.Equal(3, iterations);
    }

    [Fact]
    public void Project_LanguageDialect_DefaultsToVb6Classic_AndPersistsExtensions()
    {
        var proj = new VB6Project
        {
            Name = "LegacyProject"
        };

        // 1. Sin especificar en el .vbp, debe ser VB6 Clásico Estricto
        Assert.Equal(VbLanguageDialect.Vb6Classic, proj.Dialect);

        // 2. Asignar VB7 Office (VBA7 64-bit bridge)
        proj.Dialect = VbLanguageDialect.Vb7Office;
        Assert.Equal(VbLanguageDialect.Vb7Office, proj.Dialect);
        Assert.Equal("Vb7Office", proj.Settings["LanguageDialect"]);

        // 3. Asignar ViBa7 Moderno
        proj.Dialect = VbLanguageDialect.ViBaModern;
        Assert.Equal(VbLanguageDialect.ViBaModern, proj.Dialect);
        Assert.Equal("ViBaModern", proj.Settings["LanguageDialect"]);

        // 4. Volver a VB6 Clásico
        proj.Dialect = VbLanguageDialect.Vb6Classic;
        Assert.Equal(VbLanguageDialect.Vb6Classic, proj.Dialect);
        Assert.Equal("VB6", proj.Settings["LanguageDialect"]);
    }
}
