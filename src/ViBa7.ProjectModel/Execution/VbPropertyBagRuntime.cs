using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ViBa7.ProjectModel.Execution;

/// <summary>
/// Objeto PropertyBag ejecutable en tiempo de ejecución para UserControls (.ctl) y componentes ActiveX.
/// Implementa ILiveControl para permitir llamadas directas a PropBag.ReadProperty y PropBag.WriteProperty en código VB6,
/// así como la serialización y deserialización binaria de su contenido (Contents).
/// </summary>
public sealed class VbPropertyBagRuntime : ILiveControl
{
    private readonly Dictionary<string, VbValue> _properties = new(StringComparer.OrdinalIgnoreCase);

    public string Name => "PropBag";
    public string TypeName => "PropertyBag";
    public int Count => _properties.Count;

    public IReadOnlyDictionary<string, VbValue> Properties => _properties;

    public VbPropertyBagRuntime(IDictionary<string, string>? initialStringProperties = null)
    {
        if (initialStringProperties != null)
        {
            foreach (var kvp in initialStringProperties)
            {
                _properties[kvp.Key] = VbValue.FromString(kvp.Value);
            }
        }
    }

    public void SetPropertyDirect(string name, VbValue value)
    {
        if (!string.IsNullOrWhiteSpace(name))
        {
            _properties[name] = value;
        }
    }

    public VbValue ReadProperty(string name, VbValue? defaultValue = null)
    {
        if (string.IsNullOrWhiteSpace(name)) return defaultValue ?? VbValue.Empty;
        return _properties.TryGetValue(name, out var val) ? val : (defaultValue ?? VbValue.Empty);
    }

    public void WriteProperty(string name, VbValue value, VbValue? defaultValue = null)
    {
        if (string.IsNullOrWhiteSpace(name)) return;
        _properties[name] = value;
    }

    public VbValue GetProperty(string propertyName)
    {
        switch (propertyName.ToUpperInvariant())
        {
            case "COUNT":
                return VbValue.FromLong(_properties.Count);

            case "CONTENTS":
                return VbValue.FromObject(SerializeToBytes());

            default:
                return _properties.TryGetValue(propertyName, out var val) ? val : VbValue.Empty;
        }
    }

    public void SetProperty(string propertyName, VbValue value)
    {
        switch (propertyName.ToUpperInvariant())
        {
            case "CONTENTS":
                if (value.RawValue is byte[] bytes)
                {
                    DeserializeFromBytes(bytes);
                }
                break;

            default:
                _properties[propertyName] = value;
                break;
        }
    }

    public VbValue InvokeMethod(string methodName, VbValue[] args)
    {
        switch (methodName.ToUpperInvariant())
        {
            case "READPROPERTY":
                if (args.Length > 0)
                {
                    string propName = args[0].AsString();
                    VbValue defVal = args.Length > 1 ? args[1] : VbValue.Empty;
                    return ReadProperty(propName, defVal);
                }
                return VbValue.Empty;

            case "WRITEPROPERTY":
                if (args.Length > 1)
                {
                    string propName = args[0].AsString();
                    VbValue val = args[1];
                    VbValue defVal = args.Length > 2 ? args[2] : VbValue.Empty;
                    WriteProperty(propName, val, defVal);
                }
                return VbValue.Empty;

            case "CLEAR":
                _properties.Clear();
                return VbValue.Empty;

            default:
                return VbValue.Empty;
        }
    }

    /// <summary>
    /// Serializa los pares clave-valor a un bloque binario canónico para persistencia en .ctx / .frx.
    /// </summary>
    public byte[] SerializeToBytes()
    {
        using var ms = new MemoryStream();
        using var bw = new BinaryWriter(ms, Encoding.UTF8);

        // Cabecera: Firma "VB7P" + cantidad de entradas
        bw.Write(Encoding.ASCII.GetBytes("VB7P"));
        bw.Write(_properties.Count);

        foreach (var kvp in _properties)
        {
            bw.Write(kvp.Key);
            string strVal = kvp.Value.AsString();
            bw.Write(strVal);
        }

        return ms.ToArray();
    }

    /// <summary>
    /// Deserializa pares clave-valor desde un bloque binario persistido.
    /// </summary>
    public void DeserializeFromBytes(byte[] data)
    {
        if (data == null || data.Length < 8) return;

        using var ms = new MemoryStream(data);
        using var br = new BinaryReader(ms, Encoding.UTF8);

        string magic = Encoding.ASCII.GetString(br.ReadBytes(4));
        if (magic != "VB7P") return;

        int count = br.ReadInt32();
        for (int i = 0; i < count && ms.Position < ms.Length; i++)
        {
            string key = br.ReadString();
            string val = br.ReadString();
            _properties[key] = VbValue.FromString(val);
        }
    }
}
