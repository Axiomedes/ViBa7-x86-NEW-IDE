using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace ViBa7.ProjectModel.Execution;

public enum FlowSignal
{
    Normal,
    ExitSub,
    ExitFunction,
    ExitFor,
    ExitDo,
    EndExecution
}

/// <summary>
/// Intérprete y ejecutor de sentencias de Visual Basic 6.0 en caliente (Hot-Run).
/// </summary>
public class VbStatementInterpreter
{
    private readonly VbExecutionContext _context;

    public VbStatementInterpreter(VbExecutionContext context)
    {
        _context = context;
    }

    /// <summary>
    /// Ejecuta un bloque completo de líneas de código VB6 (cuerpo de un procedimiento o evento).
    /// </summary>
    public FlowSignal ExecuteBlock(IReadOnlyList<string> lines, string? sourceFile = null, int baseLineNumber = 1)
    {
        if (_context.State == ExecutionState.Stopped)
        {
            _context.State = ExecutionState.Running;
        }
        int index = 0;
        return ExecuteBlockInternal(lines, ref index, sourceFile, baseLineNumber);
    }

    private FlowSignal ExecuteBlockInternal(IReadOnlyList<string> lines, ref int index, string? sourceFile, int baseLineNumber)
    {
        while (index < lines.Count)
        {
            if (_context.State == ExecutionState.Stopped)
            {
                return FlowSignal.EndExecution;
            }

            int currentLine = baseLineNumber + index;
            string rawLine = lines[index].Trim();
            index++;

            // Líneas vacías o comentarios puros
            if (string.IsNullOrWhiteSpace(rawLine) || rawLine.StartsWith('\'') || rawLine.StartsWith("Rem ", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            // Comprobar breakpoints activos o modo paso a paso
            bool hasBp = _context.HasBreakpoint(sourceFile, currentLine);
            bool isStepping = _context.CurrentStepMode == StepMode.StepInto ||
                              (_context.CurrentStepMode == StepMode.StepOver && _context.CallStack.Count <= _context.StepOverFrameDepth);

            if ((hasBp || isStepping) && _context.State != ExecutionState.Stopped)
            {
                if (_context.CurrentFrame != null)
                {
                    _context.CurrentFrame.CurrentLineNumber = currentLine;
                }

                _context.TriggerBreakpoint(sourceFile ?? string.Empty, currentLine);

                if (_context.State == ExecutionState.Stopped)
                {
                    return FlowSignal.ExitSub;
                }
            }

            // Ejecutar sentencia con captura de errores según On Error
            try
            {
                var signal = ExecuteSingleStatement(rawLine, ref index, lines, sourceFile, baseLineNumber);
                if (signal != FlowSignal.Normal)
                {
                    return signal;
                }
            }
            catch (Exception ex)
            {
                var errorMode = _context.CurrentFrame?.ErrorMode ?? _context.GlobalErrorMode;

                int errNum = ex is DivideByZeroException ? 11 : 5;
                _context.Err.Raise(errNum, description: ex.Message);

                if (errorMode == OnErrorHandling.ResumeNext)
                {
                    // Ignora el error y continúa con la siguiente instrucción
                    continue;
                }
                else if (errorMode == OnErrorHandling.GoToLabel && !string.IsNullOrEmpty(_context.CurrentFrame?.ErrorHandlerLabel))
                {
                    string targetLabel = _context.CurrentFrame.ErrorHandlerLabel;
                    int labelIdx = -1;
                    for (int l = 0; l < lines.Count; l++)
                    {
                        string lineTrim = lines[l].Trim();
                        if (lineTrim.Equals(targetLabel + ":", StringComparison.OrdinalIgnoreCase) ||
                            lineTrim.Equals(targetLabel, StringComparison.OrdinalIgnoreCase))
                        {
                            labelIdx = l;
                            break;
                        }
                    }

                    if (labelIdx >= 0)
                    {
                        index = labelIdx + 1;
                        continue;
                    }
                }

                throw;
            }
        }

        return FlowSignal.Normal;
    }

    public FlowSignal ExecuteSingleStatement(string rawLine)
    {
        int dummy = 0;
        return ExecuteSingleStatement(rawLine, ref dummy);
    }

    public FlowSignal ExecuteSingleStatement(
        string rawLine,
        ref int nextLineIndex,
        IReadOnlyList<string>? enclosingLines = null,
        string? sourceFile = null,
        int baseLineNumber = 1)
    {
        string line = rawLine.Trim();

        // Eliminar comentarios al final de la línea si no están dentro de comillas
        line = StripTrailingComment(line);
        if (string.IsNullOrWhiteSpace(line)) return FlowSignal.Normal;

        // Separar múltiples instrucciones separadas por dos puntos ':'
        var statements = SplitStatements(line);
        if (statements.Count > 1)
        {
            foreach (var st in statements)
            {
                var sig = ExecuteSingleStatement(st, ref nextLineIndex, enclosingLines, sourceFile, baseLineNumber);
                if (sig != FlowSignal.Normal) return sig;
            }
            return FlowSignal.Normal;
        }

        // 1. Manejo de Errores: On Error ...
        if (line.StartsWith("On Error ", StringComparison.OrdinalIgnoreCase))
        {
            ExecuteOnError(line);
            return FlowSignal.Normal;
        }

        // 2. Sentencias Exit ...
        if (line.Equals("Exit Sub", StringComparison.OrdinalIgnoreCase)) return FlowSignal.ExitSub;
        if (line.Equals("Exit Function", StringComparison.OrdinalIgnoreCase)) return FlowSignal.ExitFunction;
        if (line.Equals("Exit For", StringComparison.OrdinalIgnoreCase)) return FlowSignal.ExitFor;
        if (line.Equals("Exit Do", StringComparison.OrdinalIgnoreCase)) return FlowSignal.ExitDo;

        // 3. Sentencia End
        if (line.Equals("End", StringComparison.OrdinalIgnoreCase))
        {
            _context.State = ExecutionState.Stopped;
            return FlowSignal.EndExecution;
        }

        // 4. Debug.Print
        if (line.StartsWith("Debug.Print", StringComparison.OrdinalIgnoreCase))
        {
            string printExpr = line["Debug.Print".Length..].Trim();
            ExecuteDebugPrint(printExpr);
            return FlowSignal.Normal;
        }

        // 5. Beep
        if (line.Equals("Beep", StringComparison.OrdinalIgnoreCase))
        {
            try { Console.Beep(); } catch { }
            return FlowSignal.Normal;
        }

        // 6. Declaraciones Dim / Static
        if (line.StartsWith("Dim ", StringComparison.OrdinalIgnoreCase) || line.StartsWith("Static ", StringComparison.OrdinalIgnoreCase))
        {
            ExecuteDim(line);
            return FlowSignal.Normal;
        }

        // 7. Bucle For ... Next
        if (line.StartsWith("For ", StringComparison.OrdinalIgnoreCase) && enclosingLines != null)
        {
            return ExecuteForLoop(line, enclosingLines, ref nextLineIndex, sourceFile, baseLineNumber);
        }

        // 8. Bucle Do ... Loop
        if (line.StartsWith("Do", StringComparison.OrdinalIgnoreCase) && enclosingLines != null)
        {
            return ExecuteDoLoop(line, enclosingLines, ref nextLineIndex, sourceFile, baseLineNumber);
        }

        // 9. Bucle While ... Wend
        if (line.StartsWith("While ", StringComparison.OrdinalIgnoreCase) && enclosingLines != null)
        {
            return ExecuteWhileLoop(line, enclosingLines, ref nextLineIndex, sourceFile, baseLineNumber);
        }

        // 10. Condicional If ... Then
        if (line.StartsWith("If ", StringComparison.OrdinalIgnoreCase))
        {
            return ExecuteIf(line, enclosingLines, ref nextLineIndex, sourceFile, baseLineNumber);
        }

        // 11. Estructura Select Case ... End Select
        if (line.StartsWith("Select Case", StringComparison.OrdinalIgnoreCase) && enclosingLines != null)
        {
            return ExecuteSelectCase(line, enclosingLines, ref nextLineIndex, sourceFile, baseLineNumber);
        }

        // 12. Invocación Call ...
        if (line.StartsWith("Call ", StringComparison.OrdinalIgnoreCase))
        {
            string callExpr = line[5..].Trim();
            VbExpressionEvaluator.Evaluate(callExpr, _context);
            return FlowSignal.Normal;
        }

        // 12. Asignación: [Let] Var/Prop = Expr o Set Obj = Expr
        if (line.StartsWith("Set ", StringComparison.OrdinalIgnoreCase))
        {
            ExecuteAssignment(line[4..].Trim());
            return FlowSignal.Normal;
        }

        if (line.StartsWith("Let ", StringComparison.OrdinalIgnoreCase))
        {
            line = line[4..].Trim();
        }

        int eqIndex = FindAssignmentEqualSign(line);
        if (eqIndex > 0)
        {
            ExecuteAssignment(line);
            return FlowSignal.Normal;
        }

        // 13. RaiseEvent
        if (line.StartsWith("RaiseEvent ", StringComparison.OrdinalIgnoreCase))
        {
            ExecuteRaiseEvent(line["RaiseEvent ".Length..].Trim());
            return FlowSignal.Normal;
        }

        // 14. DoEvents
        if (line.Equals("DoEvents", StringComparison.OrdinalIgnoreCase))
        {
            _context.DoEvents();
            return FlowSignal.Normal;
        }

        // 15. Invocación directa de método o subrutina sin Call: ej. `MsgBox "Hola"` o `MiSub 1, 2`
        ExecuteSubInvocation(line);

        return FlowSignal.Normal;
    }

    #region Ejecución de Sentencias Específicas

    private void ExecuteRaiseEvent(string rest)
    {
        if (string.IsNullOrWhiteSpace(rest)) return;

        string eventName;
        string rawArgs = string.Empty;

        int openParen = rest.IndexOf('(');
        int firstSpace = rest.IndexOf(' ');

        if (openParen > 0 && (firstSpace < 0 || openParen < firstSpace))
        {
            eventName = rest[..openParen].Trim();
            int closeParen = rest.LastIndexOf(')');
            if (closeParen > openParen)
            {
                rawArgs = rest[(openParen + 1)..closeParen].Trim();
            }
        }
        else if (firstSpace > 0)
        {
            eventName = rest[..firstSpace].Trim();
            rawArgs = rest[(firstSpace + 1)..].Trim();
        }
        else
        {
            eventName = rest.Trim();
        }

        var argValues = new List<VbValue>();
        if (!string.IsNullOrWhiteSpace(rawArgs))
        {
            var parts = SplitCommaArguments(rawArgs);
            foreach (var p in parts)
            {
                if (string.IsNullOrWhiteSpace(p))
                {
                    argValues.Add(VbValue.Empty);
                }
                else
                {
                    argValues.Add(VbExpressionEvaluator.Evaluate(p, _context));
                }
            }
        }

        _context.RaiseEvent(eventName, argValues.ToArray());
    }

    private void ExecuteDebugPrint(string expr)
    {
        if (string.IsNullOrWhiteSpace(expr))
        {
            _context.WriteDebugPrint(string.Empty);
            return;
        }

        var parts = SplitPrintArguments(expr);
        var sb = new System.Text.StringBuilder();

        foreach (var p in parts)
        {
            var val = VbExpressionEvaluator.Evaluate(p, _context);
            sb.Append(val.AsString());
            sb.Append(' ');
        }

        _context.WriteDebugPrint(sb.ToString().TrimEnd());
    }

    private void ExecuteOnError(string line)
    {
        string sub = line["On Error ".Length..].Trim();
        if (sub.Equals("Resume Next", StringComparison.OrdinalIgnoreCase))
        {
            if (_context.CurrentFrame != null)
            {
                _context.CurrentFrame.ErrorMode = OnErrorHandling.ResumeNext;
            }
            _context.GlobalErrorMode = OnErrorHandling.ResumeNext;
        }
        else if (sub.Equals("GoTo 0", StringComparison.OrdinalIgnoreCase))
        {
            if (_context.CurrentFrame != null)
            {
                _context.CurrentFrame.ErrorMode = OnErrorHandling.Standard;
                _context.CurrentFrame.ErrorHandlerLabel = null;
            }
            _context.GlobalErrorMode = OnErrorHandling.Standard;
        }
        else if (sub.StartsWith("GoTo ", StringComparison.OrdinalIgnoreCase))
        {
            string label = sub[5..].Trim();
            if (_context.CurrentFrame != null)
            {
                _context.CurrentFrame.ErrorMode = OnErrorHandling.GoToLabel;
                _context.CurrentFrame.ErrorHandlerLabel = label;
            }
        }
    }

    private void ExecuteDim(string line)
    {
        bool isStatic = line.StartsWith("Static ", StringComparison.OrdinalIgnoreCase);
        int spaceIdx = line.IndexOf(' ');
        string decls = line[(spaceIdx + 1)..].Trim();

        foreach (var decl in decls.Split(','))
        {
            string d = decl.Trim();
            int asIdx = d.IndexOf(" As ", StringComparison.OrdinalIgnoreCase);
            string varName = asIdx > 0 ? d[..asIdx].Trim() : d;

            if (isStatic && _context.CurrentFrame != null)
            {
                string staticKey = VbExecutionContext.GetStaticVarKey(_context.CurrentFrame.ProcedureName, varName);
                if (!_context.StaticVariables.ContainsKey(staticKey))
                {
                    _context.StaticVariables[staticKey] = VbValue.Empty;
                }
            }
            else
            {
                _context.SetVariable(varName, VbValue.Empty);
            }
        }
    }

    private void ExecuteAssignment(string line)
    {
        int eqIndex = FindAssignmentEqualSign(line);
        if (eqIndex <= 0) return;

        string target = line[..eqIndex].Trim();
        string expr = line[(eqIndex + 1)..].Trim();

        var value = VbExpressionEvaluator.Evaluate(expr, _context);

        // Si es asignación a propiedad de control u objeto encadenado:
        // ej. `Text1.Text = "Hola"`, `MSFlexGrid1.TextMatrix(r, c) = "X"`, `ListView1.ListItems(1).SubItems(1) = "Y"`
        int dot = FindLastTopLevelDot(target);
        if (dot > 0)
        {
            string objName = target[..dot].Trim();
            string propName = target[(dot + 1)..].Trim();

            var obj = (objName.Contains('(') || objName.Contains('.'))
                ? VbExpressionEvaluator.Evaluate(objName, _context)
                : _context.GetVariable(objName);

            if (obj.RawValue is ILiveControl ctrl)
            {
                int openParen = propName.IndexOf('(');
                int closeParen = propName.LastIndexOf(')');
                if (openParen > 0 && closeParen > openParen && closeParen == propName.Length - 1)
                {
                    string basePropName = propName[..openParen].Trim();
                    string insideArgs = propName[(openParen + 1)..closeParen].Trim();
                    var indexArgs = new List<VbValue>();
                    if (!string.IsNullOrWhiteSpace(insideArgs))
                    {
                        foreach (var argExpr in SplitCommaArguments(insideArgs))
                        {
                            indexArgs.Add(VbExpressionEvaluator.Evaluate(argExpr, _context));
                        }
                    }

                    ctrl.SetIndexedProperty(basePropName, indexArgs.ToArray(), value);
                    return;
                }

                ctrl.SetProperty(propName, value);
                return;
            }
            if (obj.RawValue is VbErrObject err)
            {
                if (propName.Equals("Number", StringComparison.OrdinalIgnoreCase)) err.Number = value.AsLong();
                else if (propName.Equals("Description", StringComparison.OrdinalIgnoreCase)) err.Description = value.AsString();
                else if (propName.Equals("Source", StringComparison.OrdinalIgnoreCase)) err.Source = value.AsString();
                return;
            }
        }

        // Asignación a variable local o global
        _context.SetVariable(target, value);
    }

    private void ExecuteSubInvocation(string line)
    {
        string subName;
        string argsStr;

        int topSpace = FindFirstTopLevelSpace(line);
        if (topSpace > 0)
        {
            subName = line[..topSpace].Trim();
            argsStr = line[(topSpace + 1)..].Trim();
        }
        else if (line.EndsWith(')') && TrySplitTrailingParenthesizedCall(line, out string? parsedSub, out string? parsedArgs))
        {
            subName = parsedSub!;
            argsStr = parsedArgs!;
        }
        else
        {
            subName = line;
            argsStr = string.Empty;
        }

        argsStr = StripOutermostParentheses(argsStr);

        var args = new List<VbValue>();
        if (!string.IsNullOrWhiteSpace(argsStr))
        {
            foreach (var a in SplitCommaArguments(argsStr))
            {
                args.Add(VbExpressionEvaluator.Evaluate(a, _context));
            }
        }

        // Comprobar funciones intrínsecas como MsgBox
        if (subName.Equals("MsgBox", StringComparison.OrdinalIgnoreCase) && args.Count > 0)
        {
            string prompt = args[0].AsString();
            int buttons = args.Count > 1 ? (int)args[1].AsLong() : 0;
            string title = args.Count > 2 ? args[2].AsString() : string.Empty;

            if (_context.MsgBoxHandler != null)
            {
                _context.MsgBoxHandler(prompt, buttons, title);
            }
            else
            {
                _context.WriteDebugPrint($"[MsgBox]: {prompt}");
            }
            return;
        }

        // Si es método de objeto/control o colección encadenada:
        // ej. `Command1.SetFocus`, `TreeView1.Nodes.Add , , "k", "Texto"`, `ListView1.ListItems(1).EnsureVisible`
        int dot = FindLastTopLevelDot(subName);
        if (dot > 0)
        {
            string objName = subName[..dot].Trim();
            string methodName = subName[(dot + 1)..].Trim();

            var obj = (objName.Contains('(') || objName.Contains('.'))
                ? VbExpressionEvaluator.Evaluate(objName, _context)
                : _context.GetVariable(objName);

            if (obj.RawValue is ILiveControl ctrl)
            {
                ctrl.InvokeMethod(methodName, args.ToArray());
                return;
            }
        }

        // Probar si existe en el runtime
        VbRuntimeFunctions.TryInvoke(subName, args.ToArray(), out _);
    }

    private static int FindLastTopLevelDot(string expr)
    {
        bool inQuotes = false;
        int parenDepth = 0;
        int lastDot = -1;

        for (int i = 0; i < expr.Length; i++)
        {
            char c = expr[i];
            if (c == '"') inQuotes = !inQuotes;
            else if (!inQuotes)
            {
                if (c == '(') parenDepth++;
                else if (c == ')') parenDepth--;
                else if (c == '.' && parenDepth == 0)
                {
                    lastDot = i;
                }
            }
        }
        return lastDot;
    }

    private static int FindFirstTopLevelSpace(string expr)
    {
        bool inQuotes = false;
        int parenDepth = 0;

        for (int i = 0; i < expr.Length; i++)
        {
            char c = expr[i];
            if (c == '"') inQuotes = !inQuotes;
            else if (!inQuotes)
            {
                if (c == '(') parenDepth++;
                else if (c == ')') parenDepth--;
                else if (c == ' ' && parenDepth == 0)
                {
                    return i;
                }
            }
        }
        return -1;
    }

    private static bool TrySplitTrailingParenthesizedCall(string expr, out string? subName, out string? argsStr)
    {
        subName = null;
        argsStr = null;
        if (!expr.EndsWith(')')) return false;

        bool inQuotes = false;
        int parenDepth = 0;
        int openParenIdx = -1;

        for (int i = expr.Length - 1; i >= 0; i--)
        {
            char c = expr[i];
            if (c == '"') inQuotes = !inQuotes;
            else if (!inQuotes)
            {
                if (c == ')') parenDepth++;
                else if (c == '(')
                {
                    parenDepth--;
                    if (parenDepth == 0)
                    {
                        openParenIdx = i;
                        break;
                    }
                }
            }
        }

        if (openParenIdx > 0)
        {
            subName = expr[..openParenIdx].Trim();
            argsStr = expr[(openParenIdx + 1)..^1].Trim();
            return true;
        }
        return false;
    }

    #endregion

    #region Estructuras de Control (If, For, Do, While)

    private FlowSignal ExecuteIf(
        string line,
        IReadOnlyList<string>? lines,
        ref int nextLineIndex,
        string? sourceFile,
        int baseLineNumber)
    {
        // 1. Detección de If de una sola línea: `If cond Then stmt [Else stmt]`
        int thenIdx = IndexOfKeyword(line, "Then");
        if (thenIdx < 0) return FlowSignal.Normal;

        string condExpr = line[3..thenIdx].Trim();
        string afterThen = line[(thenIdx + 4)..].Trim();

        bool condResult = VbExpressionEvaluator.Evaluate(condExpr, _context).AsBoolean();

        if (!string.IsNullOrWhiteSpace(afterThen))
        {
            // If de una sola línea
            int elseIdx = IndexOfKeyword(afterThen, "Else");
            if (elseIdx >= 0)
            {
                string thenStmt = afterThen[..elseIdx].Trim();
                string elseStmt = afterThen[(elseIdx + 4)..].Trim();

                if (condResult)
                {
                    return ExecuteSingleStatement(thenStmt, ref nextLineIndex, lines, sourceFile, baseLineNumber);
                }
                else
                {
                    return ExecuteSingleStatement(elseStmt, ref nextLineIndex, lines, sourceFile, baseLineNumber);
                }
            }
            else
            {
                if (condResult)
                {
                    return ExecuteSingleStatement(afterThen, ref nextLineIndex, lines, sourceFile, baseLineNumber);
                }
                return FlowSignal.Normal;
            }
        }

        // 2. If en Bloque multilínea: If cond Then ... [ElseIf ...] [Else ...] End If
        if (lines == null) return FlowSignal.Normal;

        var branches = new List<(string? ConditionExpr, List<string> Body)>();
        var currentBody = new List<string>();
        branches.Add((condExpr, currentBody));

        int depth = 1;
        while (nextLineIndex < lines.Count)
        {
            string curLine = lines[nextLineIndex++];
            string trimmed = curLine.Trim();
            string upper = trimmed.ToUpperInvariant();

            if (upper.StartsWith("IF ") && IndexOfKeyword(curLine, "THEN") >= 0 && string.IsNullOrWhiteSpace(curLine[(IndexOfKeyword(curLine, "THEN") + 4)..].Trim()))
            {
                depth++;
                currentBody.Add(curLine);
                continue;
            }

            if (upper.StartsWith("END IF") || upper == "ENDIF")
            {
                depth--;
                if (depth == 0) break;
                currentBody.Add(curLine);
                continue;
            }

            if (depth == 1 && upper.StartsWith("ELSEIF "))
            {
                int eiThen = IndexOfKeyword(trimmed, "Then");
                string eiCond = eiThen > 0 ? trimmed["ElseIf".Length..eiThen].Trim() : trimmed["ElseIf".Length..].Trim();
                currentBody = new List<string>();
                branches.Add((eiCond, currentBody));
                continue;
            }

            if (depth == 1 && (upper == "ELSE" || (upper.StartsWith("ELSE ") && !upper.StartsWith("ELSEIF"))))
            {
                currentBody = new List<string>();
                branches.Add((null, currentBody));
                continue;
            }

            currentBody.Add(curLine);
        }

        foreach (var branch in branches)
        {
            bool isMatch = branch.ConditionExpr == null || VbExpressionEvaluator.Evaluate(branch.ConditionExpr, _context).AsBoolean();
            if (isMatch)
            {
                if (branch.Body.Count > 0)
                {
                    int bIdx = 0;
                    return ExecuteBlockInternal(branch.Body, ref bIdx, sourceFile, baseLineNumber);
                }
                return FlowSignal.Normal;
            }
        }

        return FlowSignal.Normal;
    }

    private FlowSignal ExecuteSelectCase(
        string line,
        IReadOnlyList<string> lines,
        ref int nextLineIndex,
        string? sourceFile,
        int baseLineNumber)
    {
        string testExpr = line["Select Case".Length..].Trim();
        var testVal = VbExpressionEvaluator.Evaluate(testExpr, _context);

        var cases = new List<(List<string> Conds, List<string> Body, bool IsElse)>();
        List<string>? currentBody = null;
        int depth = 1;

        while (nextLineIndex < lines.Count)
        {
            string curLine = lines[nextLineIndex++];
            string trimmed = curLine.Trim();
            string upper = trimmed.ToUpperInvariant();

            if (upper.StartsWith("SELECT CASE ", StringComparison.Ordinal) || upper == "SELECT CASE")
            {
                depth++;
                currentBody?.Add(curLine);
                continue;
            }

            if (upper.StartsWith("END SELECT", StringComparison.Ordinal) || upper == "ENDSELECT")
            {
                depth--;
                if (depth == 0) break;
                currentBody?.Add(curLine);
                continue;
            }

            if (depth == 1 && (upper.StartsWith("CASE ") || upper == "CASE"))
            {
                string caseClause = trimmed["Case".Length..].Trim();
                if (caseClause.StartsWith("Else", StringComparison.OrdinalIgnoreCase))
                {
                    currentBody = new List<string>();
                    cases.Add((new List<string>(), currentBody, true));
                }
                else
                {
                    var conds = SplitCommaArguments(caseClause);
                    currentBody = new List<string>();
                    cases.Add((conds, currentBody, false));
                }
                continue;
            }

            currentBody?.Add(curLine);
        }

        List<string>? matchingBody = null;
        foreach (var c in cases)
        {
            if (c.IsElse)
            {
                matchingBody ??= c.Body;
                continue;
            }

            bool matched = false;
            foreach (var cond in c.Conds)
            {
                string condTrim = cond.Trim();
                if (string.IsNullOrEmpty(condTrim)) continue;

                if (condTrim.StartsWith("Is ", StringComparison.OrdinalIgnoreCase))
                {
                    string rest = condTrim[3..].Trim();
                    string op = "";
                    string opExpr = "";
                    if (rest.StartsWith("<=") || rest.StartsWith(">=") || rest.StartsWith("<>"))
                    {
                        op = rest[..2];
                        opExpr = rest[2..].Trim();
                    }
                    else if (rest.StartsWith("<") || rest.StartsWith(">") || rest.StartsWith("="))
                    {
                        op = rest[..1];
                        opExpr = rest[1..].Trim();
                    }

                    if (!string.IsNullOrEmpty(op))
                    {
                        var rightVal = VbExpressionEvaluator.Evaluate(opExpr, _context);
                        matched = EvaluateComparison(testVal, op, rightVal);
                        if (matched) break;
                    }
                }
                else if (Regex.IsMatch(condTrim, @"\s+To\s+", RegexOptions.IgnoreCase))
                {
                    var m = Regex.Match(condTrim, @"^(.*?)\s+To\s+(.*)$", RegexOptions.IgnoreCase);
                    if (m.Success)
                    {
                        var lowVal = VbExpressionEvaluator.Evaluate(m.Groups[1].Value.Trim(), _context);
                        var highVal = VbExpressionEvaluator.Evaluate(m.Groups[2].Value.Trim(), _context);
                        if (VbValue.CompareGreaterOrEqual(testVal, lowVal).AsBoolean() &&
                            VbValue.CompareLessOrEqual(testVal, highVal).AsBoolean())
                        {
                            matched = true;
                            break;
                        }
                    }
                }
                else
                {
                    var val = VbExpressionEvaluator.Evaluate(condTrim, _context);
                    if (VbValue.CompareEqual(testVal, val).AsBoolean())
                    {
                        matched = true;
                        break;
                    }
                }
            }

            if (matched)
            {
                matchingBody = c.Body;
                break;
            }
        }

        if (matchingBody != null && matchingBody.Count > 0)
        {
            int bIdx = 0;
            return ExecuteBlockInternal(matchingBody, ref bIdx, sourceFile, baseLineNumber);
        }

        return FlowSignal.Normal;
    }

    private static bool EvaluateComparison(VbValue left, string op, VbValue right)
    {
        return op switch
        {
            "=" => VbValue.CompareEqual(left, right).AsBoolean(),
            "<>" => VbValue.CompareNotEqual(left, right).AsBoolean(),
            "<" => VbValue.CompareLessThan(left, right).AsBoolean(),
            ">" => VbValue.CompareGreaterThan(left, right).AsBoolean(),
            "<=" => VbValue.CompareLessOrEqual(left, right).AsBoolean(),
            ">=" => VbValue.CompareGreaterOrEqual(left, right).AsBoolean(),
            _ => false
        };
    }

    private FlowSignal ExecuteForLoop(
        string line,
        IReadOnlyList<string> lines,
        ref int nextLineIndex,
        string? sourceFile,
        int baseLineNumber)
    {
        // Sintaxis: For var = start To end [Step step]
        var match = Regex.Match(line, @"^For\s+([a-zA-Z0-9_]+)\s*=\s*(.+?)\s+To\s+(.+?)(?:\s+Step\s+(.+?))?$", RegexOptions.IgnoreCase);
        if (!match.Success) return FlowSignal.Normal;

        string varName = match.Groups[1].Value;
        string startExpr = match.Groups[2].Value;
        string endExpr = match.Groups[3].Value;
        string stepExpr = match.Groups[4].Success ? match.Groups[4].Value : "1";

        int bodyStartIndex = nextLineIndex;
        var loopBody = new List<string>();
        int depth = 1;

        while (nextLineIndex < lines.Count)
        {
            string curLine = lines[nextLineIndex++];
            string upper = curLine.Trim().ToUpperInvariant();

            if (upper.StartsWith("FOR ")) depth++;
            else if (upper.StartsWith("NEXT"))
            {
                depth--;
                if (depth == 0) break;
            }

            loopBody.Add(curLine);
        }

        double startVal = VbExpressionEvaluator.Evaluate(startExpr, _context).AsDouble();
        double endVal = VbExpressionEvaluator.Evaluate(endExpr, _context).AsDouble();
        double stepVal = VbExpressionEvaluator.Evaluate(stepExpr, _context).AsDouble();

        for (double v = startVal; (stepVal > 0 ? v <= endVal : v >= endVal); v += stepVal)
        {
            _context.SetVariable(varName, VbValue.FromDouble(v));
            int bodyIdx = 0;
            var sig = ExecuteBlockInternal(loopBody, ref bodyIdx, sourceFile, baseLineNumber + bodyStartIndex);
            if (sig == FlowSignal.ExitFor) break;
            if (sig is FlowSignal.ExitSub or FlowSignal.ExitFunction or FlowSignal.EndExecution) return sig;
        }

        return FlowSignal.Normal;
    }

    private FlowSignal ExecuteDoLoop(
        string line,
        IReadOnlyList<string> lines,
        ref int nextLineIndex,
        string? sourceFile,
        int baseLineNumber)
    {
        int bodyStartIndex = nextLineIndex;
        var loopBody = new List<string>();
        int depth = 1;
        string loopLine = string.Empty;

        while (nextLineIndex < lines.Count)
        {
            string curLine = lines[nextLineIndex++];
            string upper = curLine.Trim().ToUpperInvariant();

            if (upper.StartsWith("DO")) depth++;
            else if (upper.StartsWith("LOOP"))
            {
                depth--;
                if (depth == 0)
                {
                    loopLine = curLine.Trim();
                    break;
                }
            }

            loopBody.Add(curLine);
        }

        while (true)
        {
            if (_context.State == ExecutionState.Stopped) return FlowSignal.EndExecution;

            // Condición inicial en Do While / Do Until
            if (line.StartsWith("Do While ", StringComparison.OrdinalIgnoreCase))
            {
                if (!VbExpressionEvaluator.Evaluate(line[9..], _context).AsBoolean()) break;
            }
            else if (line.StartsWith("Do Until ", StringComparison.OrdinalIgnoreCase))
            {
                if (VbExpressionEvaluator.Evaluate(line[9..], _context).AsBoolean()) break;
            }

            int bodyIdx = 0;
            var sig = ExecuteBlockInternal(loopBody, ref bodyIdx, sourceFile, baseLineNumber + bodyStartIndex);
            if (sig == FlowSignal.ExitDo) break;
            if (sig is FlowSignal.ExitSub or FlowSignal.ExitFunction or FlowSignal.EndExecution) return sig;

            // Condición final en Loop While / Loop Until
            if (loopLine.StartsWith("Loop While ", StringComparison.OrdinalIgnoreCase))
            {
                if (!VbExpressionEvaluator.Evaluate(loopLine[11..], _context).AsBoolean()) break;
            }
            else if (loopLine.StartsWith("Loop Until ", StringComparison.OrdinalIgnoreCase))
            {
                if (VbExpressionEvaluator.Evaluate(loopLine[11..], _context).AsBoolean()) break;
            }
        }

        return FlowSignal.Normal;
    }

    private FlowSignal ExecuteWhileLoop(
        string line,
        IReadOnlyList<string> lines,
        ref int nextLineIndex,
        string? sourceFile,
        int baseLineNumber)
    {
        string condExpr = line[6..].Trim();
        int bodyStartIndex = nextLineIndex;
        var loopBody = new List<string>();
        int depth = 1;

        while (nextLineIndex < lines.Count)
        {
            string curLine = lines[nextLineIndex++];
            string upper = curLine.Trim().ToUpperInvariant();

            if (upper.StartsWith("WHILE ")) depth++;
            else if (upper == "WEND")
            {
                depth--;
                if (depth == 0) break;
            }

            loopBody.Add(curLine);
        }

        while (VbExpressionEvaluator.Evaluate(condExpr, _context).AsBoolean())
        {
            if (_context.State == ExecutionState.Stopped) return FlowSignal.EndExecution;

            int bodyIdx = 0;
            var sig = ExecuteBlockInternal(loopBody, ref bodyIdx, sourceFile, baseLineNumber + bodyStartIndex);
            if (sig == FlowSignal.ExitDo || sig == FlowSignal.ExitFor) break;
            if (sig is FlowSignal.ExitSub or FlowSignal.ExitFunction or FlowSignal.EndExecution) return sig;
        }

        return FlowSignal.Normal;
    }

    #endregion

    #region Utilidades Auxiliares de Parsing

    private static string StripTrailingComment(string line)
    {
        bool inQuotes = false;
        for (int i = 0; i < line.Length; i++)
        {
            if (line[i] == '"') inQuotes = !inQuotes;
            else if (line[i] == '\'' && !inQuotes) return line[..i].Trim();
        }
        return line;
    }

    private static List<string> SplitStatements(string line)
    {
        var result = new List<string>();
        int start = 0;
        bool inQuotes = false;

        for (int i = 0; i < line.Length; i++)
        {
            if (line[i] == '"') inQuotes = !inQuotes;
            else if (line[i] == ':' && !inQuotes)
            {
                result.Add(line[start..i].Trim());
                start = i + 1;
            }
        }

        if (start < line.Length)
        {
            result.Add(line[start..].Trim());
        }

        return result;
    }

    private static int FindAssignmentEqualSign(string line)
    {
        bool inQuotes = false;
        int parenDepth = 0;

        for (int i = 0; i < line.Length; i++)
        {
            char c = line[i];
            if (c == '"') inQuotes = !inQuotes;
            else if (!inQuotes)
            {
                if (c == '(') parenDepth++;
                else if (c == ')') parenDepth--;
                else if (c == '=' && parenDepth == 0)
                {
                    // Asegurar que no sea <= o >= o ==
                    if (i > 0 && (line[i - 1] == '<' || line[i - 1] == '>')) continue;
                    return i;
                }
            }
        }
        return -1;
    }

    private static int IndexOfKeyword(string text, string keyword)
    {
        var match = Regex.Match(text, $@"\b{Regex.Escape(keyword)}\b", RegexOptions.IgnoreCase);
        return match.Success ? match.Index : -1;
    }

    private static List<string> SplitPrintArguments(string expr)
    {
        var list = new List<string>();
        int start = 0;
        bool inQuotes = false;

        for (int i = 0; i < expr.Length; i++)
        {
            if (expr[i] == '"') inQuotes = !inQuotes;
            else if ((expr[i] == ',' || expr[i] == ';') && !inQuotes)
            {
                string slice = expr[start..i].Trim();
                if (!string.IsNullOrEmpty(slice)) list.Add(slice);
                start = i + 1;
            }
        }

        if (start < expr.Length)
        {
            string slice = expr[start..].Trim();
            if (!string.IsNullOrEmpty(slice)) list.Add(slice);
        }

        return list;
    }

    private static string StripOutermostParentheses(string text)
    {
        string trimmed = text.Trim();
        if (trimmed.StartsWith('(') && trimmed.EndsWith(')'))
        {
            int depth = 0;
            bool wrapsAll = true;
            bool inQuotes = false;
            for (int i = 0; i < trimmed.Length - 1; i++)
            {
                if (trimmed[i] == '"') inQuotes = !inQuotes;
                else if (!inQuotes)
                {
                    if (trimmed[i] == '(') depth++;
                    else if (trimmed[i] == ')')
                    {
                        depth--;
                        if (depth <= 0)
                        {
                            wrapsAll = false;
                            break;
                        }
                    }
                }
            }
            if (wrapsAll && depth == 1)
            {
                return trimmed[1..^1].Trim();
            }
        }
        return trimmed;
    }

    private static List<string> SplitCommaArguments(string expr)
    {
        var list = new List<string>();
        int start = 0;
        bool inQuotes = false;
        int paren = 0;

        for (int i = 0; i < expr.Length; i++)
        {
            if (expr[i] == '"') inQuotes = !inQuotes;
            else if (!inQuotes)
            {
                if (expr[i] == '(') paren++;
                else if (expr[i] == ')') paren--;
                else if (expr[i] == ',' && paren == 0)
                {
                    list.Add(expr[start..i].Trim());
                    start = i + 1;
                }
            }
        }

        if (start <= expr.Length)
        {
            list.Add(expr[start..].Trim());
        }

        return list;
    }

    #endregion
}
