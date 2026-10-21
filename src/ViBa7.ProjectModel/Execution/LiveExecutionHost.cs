using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ViBa7.ProjectModel.Navigation;
using ViBa7.ProjectModel.Vbp;

namespace ViBa7.ProjectModel.Execution;

public record LoadedProcedure(
    string Name,
    string? ContainerName,
    List<string> Parameters,
    List<string> BodyLines,
    int StartLine,
    string? SourceFilePath = null
);

/// <summary>
/// Orquestador general del motor de ejecución en caliente (Hot-Run Host).
/// Indexa los procedimientos de los archivos del proyecto, resuelve invocaciones
/// de eventos y coordina la ejecución en vivo sin compilar a disco.
/// </summary>
public class LiveExecutionHost
{
    private readonly VbExecutionContext _context;
    private readonly VbStatementInterpreter _interpreter;
    private readonly Dictionary<string, LoadedProcedure> _procedures = new(StringComparer.OrdinalIgnoreCase);

    public VbExecutionContext Context => _context;
    public VbStatementInterpreter Interpreter => _interpreter;

    public LiveExecutionHost(VbExecutionContext? context = null)
    {
        _context = context ?? new VbExecutionContext();
        _interpreter = new VbStatementInterpreter(_context);
    }

    /// <summary>
    /// Indexa los procedimientos de un archivo o código fuente para hacerlos ejecutables.
    /// </summary>
    public void LoadSourceCode(string code, string? containerName = null, string? sourceFilePath = null)
    {
        var blocks = Advanced.VB6BlockParser.ParseBlocks(code);
        string[] allLines = code.Replace("\r\n", "\n").Split('\n');

        var procRegex = new System.Text.RegularExpressions.Regex(
            @"^\s*(?:(?:Public|Private|Friend|Static)\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+([a-zA-Z0-9_]+)\s*(?:\((.*?)\))?",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);

        foreach (var block in blocks)
        {
            if (block.BlockType is "Sub" or "Function" or "Property")
            {
                var match = procRegex.Match(block.Header);
                if (!match.Success) continue;

                string procName = match.Groups[1].Value;
                string rawParams = match.Groups[2].Success ? match.Groups[2].Value : string.Empty;

                var paramNames = new List<string>();
                if (!string.IsNullOrWhiteSpace(rawParams))
                {
                    foreach (var p in rawParams.Split(','))
                    {
                        string pTrim = p.Trim();
                        pTrim = System.Text.RegularExpressions.Regex.Replace(pTrim, @"^(?:ByVal|ByRef|Optional)\s+", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                        var asMatch = System.Text.RegularExpressions.Regex.Match(pTrim, @"\s+As\s+", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                        string pName = asMatch.Success ? pTrim[..asMatch.Index].Trim() : pTrim;
                        if (!string.IsNullOrEmpty(pName)) paramNames.Add(pName);
                    }
                }

                int start = Math.Max(0, block.StartLine - 1);
                int end = Math.Min(allLines.Length - 1, block.EndLine - 1);

                var body = new List<string>();
                for (int l = start + 1; l < end; l++)
                {
                    body.Add(allLines[l]);
                }

                var proc = new LoadedProcedure(
                    Name: procName,
                    ContainerName: containerName,
                    Parameters: paramNames,
                    BodyLines: body,
                    StartLine: block.StartLine,
                    SourceFilePath: sourceFilePath
                );

                string key = !string.IsNullOrEmpty(containerName)
                    ? $"{containerName}.{procName}"
                    : procName;

                _procedures[key] = proc;
                _procedures[procName] = proc;
            }
        }

        // 2. Indexar e inicializar variables y constantes a nivel de módulo (sección Declaraciones)
        var procedureLines = new HashSet<int>();
        foreach (var block in blocks)
        {
            if (block.BlockType is "Sub" or "Function" or "Property")
            {
                for (int l = Math.Max(0, block.StartLine - 1); l <= Math.Min(allLines.Length - 1, block.EndLine - 1); l++)
                {
                    procedureLines.Add(l);
                }
            }
        }

        ParseDeclarationsSection(allLines, procedureLines, containerName);
    }

    private void ParseDeclarationsSection(string[] allLines, HashSet<int> procedureLines, string? containerName)
    {
        for (int l = 0; l < allLines.Length; l++)
        {
            if (procedureLines.Contains(l)) continue;

            string line = allLines[l].Trim();
            if (string.IsNullOrWhiteSpace(line) || line.StartsWith('\'') || line.StartsWith("Rem ", StringComparison.OrdinalIgnoreCase))
                continue;

            if (line.StartsWith("Attribute ", StringComparison.OrdinalIgnoreCase) ||
                line.StartsWith("Option ", StringComparison.OrdinalIgnoreCase))
                continue;

            // Constantes a nivel de módulo: [Public|Private] Const NOMBRE = VALOR
            var constMatch = System.Text.RegularExpressions.Regex.Match(
                line,
                @"^(?:(?:Public|Private|Global)\s+)?Const\s+([a-zA-Z0-9_]+)\s*(?:As\s+[a-zA-Z0-9_]+)?\s*=\s*(.+)$",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);

            if (constMatch.Success)
            {
                string constName = constMatch.Groups[1].Value.Trim();
                string constExpr = constMatch.Groups[2].Value.Trim();
                var val = VbExpressionEvaluator.Evaluate(constExpr, _context);

                _context.ModuleVariables[constName] = val;
                if (!string.IsNullOrEmpty(containerName))
                {
                    _context.ModuleVariables[$"{containerName}.{constName}"] = val;
                }
                continue;
            }

            // Variables a nivel de módulo: (Dim|Private|Public|Global) Var1 [As Type], Var2 ...
            var varDeclMatch = System.Text.RegularExpressions.Regex.Match(
                line,
                @"^(?:Dim|Private|Public|Global)\s+(?!Sub\b|Function\b|Property\b|Type\b|Enum\b|Declare\b|Event\b)(.+)$",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);

            if (varDeclMatch.Success)
            {
                string decls = varDeclMatch.Groups[1].Value.Trim();
                foreach (var d in decls.Split(','))
                {
                    string dTrim = d.Trim();
                    int asIdx = dTrim.IndexOf(" As ", StringComparison.OrdinalIgnoreCase);
                    string vName = asIdx > 0 ? dTrim[..asIdx].Trim() : dTrim;
                    vName = vName.TrimEnd('(', ')');
                    if (string.IsNullOrEmpty(vName)) continue;

                    _context.ModuleVariables[vName] = VbValue.Empty;
                    if (!string.IsNullOrEmpty(containerName))
                    {
                        _context.ModuleVariables[$"{containerName}.{vName}"] = VbValue.Empty;
                    }
                }
            }
        }
    }

    /// <summary>
    /// Invoca un procedimiento por nombre (ej. "Form_Load" o "Command1_Click").
    /// </summary>
    public FlowSignal InvokeProcedure(string procedureName, string? containerName = null, params VbValue[] args)
    {
        string lookupKey = !string.IsNullOrEmpty(containerName)
            ? $"{containerName}.{procedureName}"
            : procedureName;

        if (!_procedures.TryGetValue(lookupKey, out var proc))
        {
            if (!_procedures.TryGetValue(procedureName, out proc))
            {
                return FlowSignal.Normal; // Si no existe el manejador de evento, no hace nada
            }
        }

        if (_context.State == ExecutionState.Stopped)
        {
            _context.State = ExecutionState.Running;
        }

        _context.PushFrame(proc.Name, proc.ContainerName);

        try
        {
            // Asignar argumentos a los parámetros del procedimiento directamente en el marco local
            for (int i = 0; i < proc.Parameters.Count; i++)
            {
                string paramName = proc.Parameters[i];
                var val = i < args.Length ? args[i] : VbValue.Empty;
                if (_context.CurrentFrame != null)
                {
                    _context.CurrentFrame.LocalVariables[paramName] = val;
                }
                else
                {
                    _context.SetVariable(paramName, val);
                }
            }

            return _interpreter.ExecuteBlock(proc.BodyLines, proc.SourceFilePath ?? containerName, proc.StartLine + 1);
        }
        finally
        {
            if (_context.CurrentFrame != null)
            {
                for (int i = 0; i < proc.Parameters.Count && i < args.Length; i++)
                {
                    string paramName = proc.Parameters[i];
                    if (_context.CurrentFrame.LocalVariables.TryGetValue(paramName, out var updatedVal))
                    {
                        args[i] = updatedVal;
                    }
                }
            }
            _context.PopFrame();
        }
    }

    /// <summary>
    /// Evalúa una expresión interactiva (útil para la Ventana Inmediata con `?` o comandos directos).
    /// </summary>
    public VbValue EvaluateImmediate(string input)
    {
        string trimmed = input.Trim();
        if (string.IsNullOrWhiteSpace(trimmed)) return VbValue.Empty;

        // Comando de impresión: `? expr` o `Print expr`
        if (trimmed.StartsWith('?'))
        {
            string expr = trimmed[1..].Trim();
            return VbExpressionEvaluator.Evaluate(expr, _context);
        }

        if (trimmed.StartsWith("Print ", StringComparison.OrdinalIgnoreCase))
        {
            string expr = trimmed[6..].Trim();
            return VbExpressionEvaluator.Evaluate(expr, _context);
        }

        // Ejecutar como sentencia
        _interpreter.ExecuteSingleStatement(trimmed);
        return VbValue.Empty;
    }

    public void Reset()
    {
        _procedures.Clear();
        _context.Clear();
    }
}
