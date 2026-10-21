using System;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;

namespace ViBa7.ProjectModel.Resources;

public enum VbImageFormat
{
    Unknown,
    Bmp,
    Ico,
    Jpeg,
    Gif,
    Png,
    WebP
}

/// <summary>
/// Gestiona la lectura y escritura de recursos binarios en archivos .frx (y .ctx / .pgx)
/// de formularios y controles de Visual Basic 6.0, manteniendo el formato binario canónico
/// de OLE StdPicture (preámbulo 0x0000746C + longitud + bytes) y compatibilidad total.
/// </summary>
public static class FrxResourceFile
{
    public const uint OlePictureSignature = 0x0000746C; // "lt\0\0" en little-endian

    private static readonly Regex FrxRefRegex = new(
        @"^""?(?:(?<file>[^""]+?):)?(?<offset>[0-9A-Fa-f]{1,8})""?$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    /// <summary>
    /// Parsea una referencia clásica de recurso FRX (ej. `"Form1.frx":08CA`, `Form1.frx:08CA` o `08CA`).
    /// </summary>
    public static bool ParseReference(string? reference, out string frxFileName, out int offset)
    {
        frxFileName = string.Empty;
        offset = 0;

        if (string.IsNullOrWhiteSpace(reference))
        {
            return false;
        }

        string clean = reference.Trim();
        int colonIdx = clean.LastIndexOf(':');
        if (colonIdx >= 0)
        {
            string filePart = clean.Substring(0, colonIdx).Trim().Trim('"');
            string hexPart = clean.Substring(colonIdx + 1).Trim().Trim('"');
            frxFileName = filePart;
            if (int.TryParse(hexPart, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int parsedOffset))
            {
                offset = parsedOffset;
                return true;
            }
        }
        else
        {
            string hexPart = clean.Trim().Trim('"');
            if (int.TryParse(hexPart, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int parsedOffset))
            {
                frxFileName = string.Empty;
                offset = parsedOffset;
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Formatea una referencia en sintaxis estándar de VB6 (ej. `"Form1.frx":0000`).
    /// </summary>
    public static string FormatReference(string frxFileName, int offset)
    {
        string hex = offset.ToString("X4", CultureInfo.InvariantCulture);
        return string.IsNullOrEmpty(frxFileName) ? hex : $"\"{frxFileName}\":{hex}";
    }

    /// <summary>
    /// Lee la carga útil binaria de una imagen o recurso desde un archivo .frx en el desplazamiento especificado.
    /// Soporta preámbulos OLE StdPicture (8 y 12 bytes) e imágenes binarias directas.
    /// </summary>
    public static byte[]? ReadResource(string frxFilePath, int offset)
    {
        if (!File.Exists(frxFilePath))
        {
            return null;
        }

        try
        {
            byte[] allBytes = File.ReadAllBytes(frxFilePath);
            return ReadResource(allBytes, offset);
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Lee la carga útil de un recurso a partir de un arreglo de bytes en el desplazamiento indicado.
    /// </summary>
    public static byte[]? ReadResource(byte[] frxData, int offset)
    {
        if (frxData == null || offset < 0 || offset >= frxData.Length)
        {
            return null;
        }

        int available = frxData.Length - offset;
        if (available < 4)
        {
            return null;
        }

        // Caso 1a: Estructura canónica de VB6 con OuterLength de 4 bytes antes de OlePictureSignature
        // Formato binario: [OuterLength 4B][Magic 0x0000746C 4B][PayloadLength 4B][Payload...]
        if (available >= 12 && BitConverter.ToUInt32(frxData, offset + 4) == OlePictureSignature)
        {
            uint payloadLength = BitConverter.ToUInt32(frxData, offset + 8);
            if (payloadLength > 0 && offset + 12 + payloadLength <= frxData.Length)
            {
                byte[] result = new byte[payloadLength];
                Buffer.BlockCopy(frxData, offset + 12, result, 0, (int)payloadLength);
                return result;
            }
        }

        uint magic = BitConverter.ToUInt32(frxData, offset);

        // Caso 1b: Estándar OLE StdPicture sin OuterLength (Magic 0x0000746C en offset directo)
        if (magic == OlePictureSignature && available >= 8)
        {
            uint payloadLength = BitConverter.ToUInt32(frxData, offset + 4);

            // Variante estándar de 8 bytes de cabecera: [Magic 4B][Length 4B][Payload...]
            if (payloadLength > 0 && offset + 8 + payloadLength <= frxData.Length)
            {
                // Verificar si los datos comienzan en offset + 8
                var format8 = DetectImageFormat(new ReadOnlySpan<byte>(frxData, offset + 8, Math.Min(16, (int)payloadLength)));
                if (format8 != VbImageFormat.Unknown)
                {
                    byte[] result = new byte[payloadLength];
                    Buffer.BlockCopy(frxData, offset + 8, result, 0, (int)payloadLength);
                    return result;
                }

                // Variante de 12 bytes de cabecera: [Magic 4B][Type 4B][Length 4B][Payload...]
                if (available >= 12)
                {
                    uint payloadLength12 = BitConverter.ToUInt32(frxData, offset + 8);
                    if (payloadLength12 > 0 && offset + 12 + payloadLength12 <= frxData.Length)
                    {
                        var format12 = DetectImageFormat(new ReadOnlySpan<byte>(frxData, offset + 12, Math.Min(16, (int)payloadLength12)));
                        if (format12 != VbImageFormat.Unknown)
                        {
                            byte[] result12 = new byte[payloadLength12];
                            Buffer.BlockCopy(frxData, offset + 12, result12, 0, (int)payloadLength12);
                            return result12;
                        }
                    }
                }

                // Si no se reconoció una cabecera gráfica específica, extraer payload estándar
                byte[] rawResult = new byte[payloadLength];
                Buffer.BlockCopy(frxData, offset + 8, rawResult, 0, (int)payloadLength);
                return rawResult;
            }
        }

        // Caso 1c: [OuterLength 4B][Payload...] donde el payload comienza directamente con una firma de imagen
        if (available >= 8)
        {
            uint outerLength = BitConverter.ToUInt32(frxData, offset);
            if (outerLength > 0 && outerLength <= available - 4)
            {
                var directFormatPrefix = DetectImageFormat(new ReadOnlySpan<byte>(frxData, offset + 4, Math.Min(16, (int)outerLength)));
                if (directFormatPrefix != VbImageFormat.Unknown)
                {
                    byte[] result = new byte[outerLength];
                    Buffer.BlockCopy(frxData, offset + 4, result, 0, (int)outerLength);
                    return result;
                }
            }
        }

        // Caso 2: Carga binaria directa sin cabecera OLE
        var directFormat = DetectImageFormat(new ReadOnlySpan<byte>(frxData, offset, Math.Min(16, available)));
        if (directFormat != VbImageFormat.Unknown)
        {
            int length = GetDirectImageLength(frxData, offset, directFormat);
            if (length > 0 && offset + length <= frxData.Length)
            {
                byte[] result = new byte[length];
                Buffer.BlockCopy(frxData, offset, result, 0, length);
                return result;
            }

            // Fallback: leer hasta el final
            byte[] rest = new byte[available];
            Buffer.BlockCopy(frxData, offset, rest, 0, available);
            return rest;
        }

        return null;
    }

    /// <summary>
    /// Lee un bloque binario arbitrario (stream de persistencia ActiveX IPersistStreamInit, PropertyBag binario o recurso)
    /// desde un archivo .frx o .ctx en el desplazamiento especificado.
    /// </summary>
    public static byte[]? ReadRawBlock(string frxFilePath, int offset)
    {
        if (!File.Exists(frxFilePath))
        {
            return null;
        }

        try
        {
            byte[] allBytes = File.ReadAllBytes(frxFilePath);
            return ReadRawBlock(allBytes, offset);
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Lee un bloque binario arbitrario desde un arreglo de bytes .frx/.ctx en el desplazamiento indicado.
    /// Soporta preámbulos OLE StdPicture (8 y 12 bytes) y bloques binarios de OCX prefijados con longitud de 4 bytes.
    /// </summary>
    public static byte[]? ReadRawBlock(byte[] frxData, int offset)
    {
        if (frxData == null || offset < 0 || offset >= frxData.Length)
        {
            return null;
        }

        // Primero intentar decodificación estructurada si es un recurso OLE StdPicture
        int available = frxData.Length - offset;
        if (available >= 8)
        {
            uint magic0 = BitConverter.ToUInt32(frxData, offset);
            uint magic4 = available >= 12 ? BitConverter.ToUInt32(frxData, offset + 4) : 0;
            if (magic0 == OlePictureSignature || magic4 == OlePictureSignature)
            {
                var picBytes = ReadResource(frxData, offset);
                if (picBytes != null) return picBytes;
            }
        }

        if (available < 4)
        {
            byte[] tail = new byte[available];
            Buffer.BlockCopy(frxData, offset, tail, 0, available);
            return tail;
        }

        // Bloque binario estándar de VB6: [Longitud 4 bytes little-endian][Payload...]
        uint blockLength = BitConverter.ToUInt32(frxData, offset);
        if (blockLength == 0)
        {
            return Array.Empty<byte>();
        }

        if (blockLength <= (uint)(available - 4))
        {
            byte[] payload = new byte[blockLength];
            Buffer.BlockCopy(frxData, offset + 4, payload, 0, (int)blockLength);
            return payload;
        }

        // Fallback: devolver bytes restantes desde offset
        byte[] raw = new byte[available];
        Buffer.BlockCopy(frxData, offset, raw, 0, available);
        return raw;
    }

    /// <summary>
    /// Agrega un bloque binario opaco (ej. estado IPersistStreamInit de un OCX) al archivo .frx/.ctx
    /// precedido por su longitud de 4 bytes (formato canónico VB6) y retorna el desplazamiento en bytes.
    /// </summary>
    public static int AppendRawBlock(string frxFilePath, byte[] rawData)
    {
        ArgumentNullException.ThrowIfNull(rawData);

        string? dir = Path.GetDirectoryName(frxFilePath);
        if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
        {
            Directory.CreateDirectory(dir);
        }

        using var stream = new FileStream(frxFilePath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.Read);
        int offset = (int)stream.Length;
        stream.Seek(offset, SeekOrigin.Begin);

        using var writer = new BinaryWriter(stream);
        writer.Write((uint)rawData.Length);
        writer.Write(rawData);
        writer.Flush();

        return offset;
    }

    /// <summary>
    /// Agrega una imagen o recurso binario al archivo .frx correspondiente utilizando
    /// la estructura canónica de VB6 (0x0000746C + longitud + bytes) y retorna el desplazamiento en bytes.
    /// Si el archivo no existe, lo crea automáticamente.
    /// </summary>
    public static int AppendResource(string frxFilePath, byte[] resourceData)
    {
        ArgumentNullException.ThrowIfNull(resourceData);

        string? dir = Path.GetDirectoryName(frxFilePath);
        if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
        {
            Directory.CreateDirectory(dir);
        }

        using var stream = new FileStream(frxFilePath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.Read);
        int offset = (int)stream.Length;
        stream.Seek(offset, SeekOrigin.Begin);

        using var writer = new BinaryWriter(stream);
        // Preámbulo canónico OLE StdPicture
        writer.Write(OlePictureSignature);
        // Longitud en bytes de la imagen
        writer.Write((uint)resourceData.Length);
        // Bytes de la imagen (WEBP, PNG, JPG, ICO, BMP, etc.)
        writer.Write(resourceData);
        writer.Flush();

        return offset;
    }

    /// <summary>
    /// Sincroniza y asegura que todos los recursos gráficos (PictureData, DownPictureData, DisabledPictureData, ToolboxBitmapData)
    /// y blobs binarios de estado ActiveX (ActiveXStateBlob y BinaryProperties) presentes en memoria
    /// se guarden en el archivo .frx (o .ctx para UserControls), preservando intactos los offsets originales existentes.
    /// </summary>
    public static void EnsureFrxResourcesSaved(Designers.FormDesignModel model, string frmFilePath)
    {
        if (model == null || string.IsNullOrWhiteSpace(frmFilePath)) return;

        bool isCtl = model.IsUserControl || frmFilePath.EndsWith(".ctl", StringComparison.OrdinalIgnoreCase);
        string binaryExt = isCtl ? ".ctx" : ".frx";
        string frxPath = Path.ChangeExtension(Path.GetFullPath(frmFilePath), binaryExt);
        string frxFileName = Path.GetFileName(frxPath);

        bool hasFormPic = model.PictureData != null && model.PictureData.Length > 0;
        bool hasToolboxBmp = model.ToolboxBitmapData != null && model.ToolboxBitmapData.Length > 0;
        var allControls = model.GetAllControls();

        bool hasAnyControlBinary = allControls.Any(c =>
            (c.PictureData != null && c.PictureData.Length > 0) ||
            (c.DownPictureData != null && c.DownPictureData.Length > 0) ||
            (c.DisabledPictureData != null && c.DisabledPictureData.Length > 0) ||
            (c.ActiveXStateBlob != null && c.ActiveXStateBlob.Length > 0) ||
            c.BinaryProperties.Count > 0);

        bool hasOriginalBytes = model.OriginalFrxBytes != null && model.OriginalFrxBytes.Length > 0;

        if (!hasFormPic && !hasToolboxBmp && !hasAnyControlBinary && !hasOriginalBytes)
        {
            return;
        }

        string? dir = Path.GetDirectoryName(frxPath);
        if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
        {
            Directory.CreateDirectory(dir);
        }

        using var ms = new MemoryStream();
        if (hasOriginalBytes)
        {
            ms.Write(model.OriginalFrxBytes!, 0, model.OriginalFrxBytes!.Length);
        }

        using var writer = new BinaryWriter(ms);

        string EnsurePictureWritten(string? currentRef, byte[] picData)
        {
            if (hasOriginalBytes &&
                !string.IsNullOrWhiteSpace(currentRef) &&
                ParseReference(currentRef, out _, out int existingOffset))
            {
                var existingPayload = ReadResource(model.OriginalFrxBytes!, existingOffset);
                if (existingPayload != null && existingPayload.AsSpan().SequenceEqual(picData))
                {
                    return FormatReference(frxFileName, existingOffset);
                }
            }

            int offset = (int)ms.Position;
            writer.Write(OlePictureSignature);
            writer.Write((uint)picData.Length);
            writer.Write(picData);
            return FormatReference(frxFileName, offset);
        }

        string EnsureRawBlobWritten(string? currentRef, byte[] blobData, bool forceRewrite)
        {
            if (!forceRewrite &&
                hasOriginalBytes &&
                !string.IsNullOrWhiteSpace(currentRef) &&
                ParseReference(currentRef, out _, out int existingOffset))
            {
                var existingPayload = ReadRawBlock(model.OriginalFrxBytes!, existingOffset);
                if (existingPayload != null && existingPayload.AsSpan().SequenceEqual(blobData))
                {
                    return FormatReference(frxFileName, existingOffset);
                }
            }

            int offset = (int)ms.Position;
            writer.Write((uint)blobData.Length);
            writer.Write(blobData);
            return FormatReference(frxFileName, offset);
        }

        if (hasFormPic)
        {
            model.Picture = EnsurePictureWritten(model.Picture, model.PictureData!);
        }

        if (hasToolboxBmp)
        {
            model.ToolboxBitmap = EnsurePictureWritten(model.ToolboxBitmap, model.ToolboxBitmapData!);
        }

        foreach (var ctrl in allControls)
        {
            if (ctrl.PictureData != null && ctrl.PictureData.Length > 0)
            {
                ctrl.Picture = EnsurePictureWritten(ctrl.Picture, ctrl.PictureData);
            }
            if (ctrl.DownPictureData != null && ctrl.DownPictureData.Length > 0)
            {
                ctrl.DownPicture = EnsurePictureWritten(ctrl.DownPicture, ctrl.DownPictureData);
            }
            if (ctrl.DisabledPictureData != null && ctrl.DisabledPictureData.Length > 0)
            {
                ctrl.DisabledPicture = EnsurePictureWritten(ctrl.DisabledPicture, ctrl.DisabledPictureData);
            }

            // Persistir blob de estado de instancia ActiveX (IPersistStreamInit / IPersistStream)
            if (ctrl.ActiveXStateBlob != null && ctrl.ActiveXStateBlob.Length > 0)
            {
                string statePropName = !string.IsNullOrWhiteSpace(ctrl.ActiveXStatePropertyName)
                    ? ctrl.ActiveXStatePropertyName!
                    : "Object.zyx";

                string? existingRef = ctrl.GetCustomProperty(statePropName);
                string newRef = EnsureRawBlobWritten(existingRef, ctrl.ActiveXStateBlob, ctrl.HasModifiedActiveXState);
                ctrl.ActiveXStatePropertyName = statePropName;
                ctrl.SetCustomProperty(statePropName, newRef);
                ctrl.BinaryProperties[statePropName] = ctrl.ActiveXStateBlob;
                ctrl.HasModifiedActiveXState = false;
            }

            // Persistir propiedades binarias individuales del PropertyBag o del .frm
            foreach (var binKvp in ctrl.BinaryProperties)
            {
                if (string.Equals(binKvp.Key, ctrl.ActiveXStatePropertyName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                string? existingPropVal = ctrl.GetCustomProperty(binKvp.Key);
                string newRef = EnsureRawBlobWritten(existingPropVal, binKvp.Value, forceRewrite: false);
                ctrl.SetCustomProperty(binKvp.Key, newRef);
            }
        }

        writer.Flush();
        byte[] finalFrxBytes = ms.ToArray();
        File.WriteAllBytes(frxPath, finalFrxBytes);
        model.OriginalFrxBytes = finalFrxBytes;
    }

    /// <summary>
    /// Detecta el formato gráfico a partir de las firmas mágicas de cabecera.
    /// </summary>
    public static VbImageFormat DetectImageFormat(ReadOnlySpan<byte> header)
    {
        if (header.Length < 2) return VbImageFormat.Unknown;

        // Windows Bitmap (BM)
        if (header[0] == 0x42 && header[1] == 0x4D)
        {
            return VbImageFormat.Bmp;
        }

        // Icono Windows / Cursor
        if (header.Length >= 4 && header[0] == 0x00 && header[1] == 0x00 &&
            (header[2] == 0x01 || header[2] == 0x02) && header[3] == 0x00)
        {
            return VbImageFormat.Ico;
        }

        // JPEG (0xFF, 0xD8, 0xFF)
        if (header.Length >= 3 && header[0] == 0xFF && header[1] == 0xD8 && header[2] == 0xFF)
        {
            return VbImageFormat.Jpeg;
        }

        // GIF (GIF87a o GIF89a)
        if (header.Length >= 4 && header[0] == 0x47 && header[1] == 0x49 && header[2] == 0x46 && header[3] == 0x38)
        {
            return VbImageFormat.Gif;
        }

        // PNG (0x89, 'P', 'N', 'G')
        if (header.Length >= 4 && header[0] == 0x89 && header[1] == 0x50 && header[2] == 0x4E && header[3] == 0x47)
        {
            return VbImageFormat.Png;
        }

        // WebP (RIFF .... WEBP)
        if (header.Length >= 12 &&
            header[0] == 0x52 && header[1] == 0x49 && header[2] == 0x46 && header[3] == 0x46 &&
            header[8] == 0x57 && header[9] == 0x45 && header[10] == 0x42 && header[11] == 0x50)
        {
            return VbImageFormat.WebP;
        }

        return VbImageFormat.Unknown;
    }

    private static int GetDirectImageLength(byte[] data, int offset, VbImageFormat format)
    {
        int available = data.Length - offset;

        switch (format)
        {
            case VbImageFormat.Bmp:
                if (available >= 6)
                {
                    int bmpSize = BitConverter.ToInt32(data, offset + 2);
                    if (bmpSize > 0 && bmpSize <= available) return bmpSize;
                }
                break;

            case VbImageFormat.WebP:
                if (available >= 8)
                {
                    uint riffSize = BitConverter.ToUInt32(data, offset + 4);
                    int totalSize = (int)riffSize + 8;
                    if (totalSize > 0 && totalSize <= available) return totalSize;
                }
                break;
        }

    return available;
    }

    /// <summary>
    /// Retorna una descripción legible del formato de imagen para la ventana de propiedades.
    /// </summary>
    public static string GetFormatDescription(byte[]? bytes, string? reference = null)
    {
        if (bytes == null || bytes.Length == 0)
        {
            if (!string.IsNullOrEmpty(reference))
            {
                return "(Picture)";
            }
            return "(None)";
        }

        var format = DetectImageFormat(bytes);
        return format switch
        {
            VbImageFormat.Bmp => "(Bitmap)",
            VbImageFormat.Ico => "(Icon)",
            VbImageFormat.Png => "(PNG)",
            VbImageFormat.WebP => "(WebP)",
            VbImageFormat.Jpeg => "(JPEG)",
            VbImageFormat.Gif => "(GIF)",
            _ => "(Picture)"
        };
    }
}
