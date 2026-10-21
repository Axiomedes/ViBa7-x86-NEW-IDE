using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using ViBa7.ProjectModel.Designers;

namespace ViBa7.IDE.Dialogs;

public partial class FontPickerDialog : Window
{
    private static readonly double[] StandardSizes =
    {
        8.0, 9.0, 10.0, 11.0, 12.0, 14.0, 16.0, 18.0, 20.0, 22.0, 24.0, 26.0, 28.0, 36.0, 48.0, 72.0
    };

    private readonly List<string> _allFontFamilies = new();
    private bool _updatingInternally;

    public FormFontDescription ResultFont { get; private set; }

    public FontPickerDialog(FormFontDescription? initialFont = null)
    {
        InitializeComponent();

        ResultFont = initialFont?.Clone() ?? new FormFontDescription();

        // 1. Cargar tipografías instaladas en el sistema operativo
        try
        {
            _allFontFamilies = Fonts.SystemFontFamilies
                .Select(f => f.Source)
                .Distinct()
                .OrderBy(name => name, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }
        catch
        {
            _allFontFamilies = new List<string>
            {
                "Arial", "Courier New", "MS Sans Serif", "Microsoft Sans Serif",
                "Segoe UI", "Tahoma", "Times New Roman", "Verdana"
            };
        }

        LstFontFamilies.ItemsSource = _allFontFamilies;

        // 2. Cargar tamaños estándar
        LstFontSizes.ItemsSource = StandardSizes.Select(s => s.ToString("0.##", CultureInfo.InvariantCulture)).ToList();

        // 3. Aplicar estado inicial
        ApplyInitialState(ResultFont);
    }

    public static bool ShowDialog(Window? owner, FormFontDescription initialFont, out FormFontDescription selectedFont)
    {
        var dlg = new FontPickerDialog(initialFont);
        if (owner != null)
        {
            dlg.Owner = owner;
        }

        if (dlg.ShowDialog() == true)
        {
            selectedFont = dlg.ResultFont;
            return true;
        }

        selectedFont = initialFont;
        return false;
    }

    private void ApplyInitialState(FormFontDescription font)
    {
        _updatingInternally = true;
        try
        {
            // Familia
            string familyName = !string.IsNullOrWhiteSpace(font.Name) ? font.Name : "MS Sans Serif";
            TxtFontFamily.Text = familyName;
            var match = _allFontFamilies.FirstOrDefault(f => string.Equals(f, familyName, StringComparison.OrdinalIgnoreCase))
                     ?? _allFontFamilies.FirstOrDefault(f => f.StartsWith(familyName, StringComparison.OrdinalIgnoreCase));
            if (match != null)
            {
                LstFontFamilies.SelectedItem = match;
                LstFontFamilies.ScrollIntoView(match);
            }

            // Estilo
            int styleIndex = 0;
            if (font.IsBold && font.Italic) styleIndex = 3;
            else if (font.IsBold) styleIndex = 2;
            else if (font.Italic) styleIndex = 1;

            LstFontStyles.SelectedIndex = styleIndex;
            TxtFontStyle.Text = ((ListBoxItem)LstFontStyles.Items[styleIndex]).Content?.ToString() ?? "Normal";

            // Tamaño
            double sz = font.Size > 0 ? font.Size : 8.25;
            string sizeStr = sz.ToString("0.##", CultureInfo.InvariantCulture);
            TxtFontSize.Text = sizeStr;

            var matchingSize = StandardSizes.FirstOrDefault(s => Math.Abs(s - sz) < 0.05);
            if (matchingSize > 0)
            {
                LstFontSizes.SelectedItem = matchingSize.ToString("0.##", CultureInfo.InvariantCulture);
                LstFontSizes.ScrollIntoView(LstFontSizes.SelectedItem);
            }

            // Efectos
            ChkUnderline.IsChecked = font.Underline;
            ChkStrikethrough.IsChecked = font.Strikethrough;
        }
        finally
        {
            _updatingInternally = false;
        }

        UpdatePreview();
    }

    private void TxtFontFamily_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_updatingInternally) return;

        string query = TxtFontFamily.Text.Trim();
        var match = _allFontFamilies.FirstOrDefault(f => f.StartsWith(query, StringComparison.OrdinalIgnoreCase))
                 ?? _allFontFamilies.FirstOrDefault(f => f.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0);

        if (match != null)
        {
            _updatingInternally = true;
            LstFontFamilies.SelectedItem = match;
            LstFontFamilies.ScrollIntoView(match);
            _updatingInternally = false;
        }

        UpdatePreview();
    }

    private void LstFontFamilies_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_updatingInternally) return;
        if (LstFontFamilies.SelectedItem is string selectedFamily)
        {
            _updatingInternally = true;
            TxtFontFamily.Text = selectedFamily;
            _updatingInternally = false;
            UpdatePreview();
        }
    }

    private void LstFontStyles_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_updatingInternally) return;
        if (LstFontStyles.SelectedItem is ListBoxItem item)
        {
            TxtFontStyle.Text = item.Content?.ToString() ?? "Normal";
            UpdatePreview();
        }
    }

    private void TxtFontSize_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_updatingInternally) return;

        string text = TxtFontSize.Text.Trim();
        if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed) ||
            double.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out parsed))
        {
            _updatingInternally = true;
            var match = LstFontSizes.Items.Cast<string>().FirstOrDefault(s => double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out double d) && Math.Abs(d - parsed) < 0.05);
            if (match != null)
            {
                LstFontSizes.SelectedItem = match;
                LstFontSizes.ScrollIntoView(match);
            }
            else
            {
                LstFontSizes.SelectedItem = null;
            }
            _updatingInternally = false;
        }

        UpdatePreview();
    }

    private void LstFontSizes_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_updatingInternally) return;
        if (LstFontSizes.SelectedItem is string selectedSize)
        {
            _updatingInternally = true;
            TxtFontSize.Text = selectedSize;
            _updatingInternally = false;
            UpdatePreview();
        }
    }

    private void Effect_Changed(object sender, RoutedEventArgs e)
    {
        if (_updatingInternally) return;
        UpdatePreview();
    }

    private void UpdatePreview()
    {
        if (TxtPreview == null) return;

        string familyName = !string.IsNullOrWhiteSpace(TxtFontFamily.Text) ? TxtFontFamily.Text.Trim() : "MS Sans Serif";
        TxtPreview.FontFamily = new FontFamily(familyName);

        int styleIdx = LstFontStyles.SelectedIndex;
        bool isBold = styleIdx is 2 or 3;
        bool isItalic = styleIdx is 1 or 3;

        TxtPreview.FontWeight = isBold ? FontWeights.Bold : FontWeights.Normal;
        TxtPreview.FontStyle = isItalic ? FontStyles.Italic : FontStyles.Normal;

        double size = 11.0;
        string szText = TxtFontSize?.Text?.Trim() ?? "";
        if (double.TryParse(szText, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed) ||
            double.TryParse(szText, NumberStyles.Float, CultureInfo.CurrentCulture, out parsed))
        {
            size = Math.Clamp(parsed, 6.0, 96.0);
        }

        // Convertir de puntos tipográficos a DIP de WPF (1 pt = 96/72 DIP = 1.333 DIP)
        TxtPreview.FontSize = Math.Max(8.0, size * (96.0 / 72.0));

        var decorations = new TextDecorationCollection();
        if (ChkUnderline?.IsChecked == true)
        {
            decorations.Add(TextDecorations.Underline[0]);
        }
        if (ChkStrikethrough?.IsChecked == true)
        {
            decorations.Add(TextDecorations.Strikethrough[0]);
        }
        TxtPreview.TextDecorations = decorations;
    }

    private void BtnOk_Click(object sender, RoutedEventArgs e)
    {
        string family = !string.IsNullOrWhiteSpace(TxtFontFamily.Text) ? TxtFontFamily.Text.Trim() : "MS Sans Serif";

        double size = 8.25;
        string szText = TxtFontSize.Text.Trim();
        if (double.TryParse(szText, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed) ||
            double.TryParse(szText, NumberStyles.Float, CultureInfo.CurrentCulture, out parsed))
        {
            size = Math.Clamp(parsed, 1.0, 2048.0);
        }

        int styleIdx = LstFontStyles.SelectedIndex;
        bool isBold = styleIdx is 2 or 3;
        bool isItalic = styleIdx is 1 or 3;

        ResultFont = new FormFontDescription
        {
            Name = family,
            Size = size,
            Weight = isBold ? 700 : 400,
            Italic = isItalic,
            Underline = ChkUnderline.IsChecked == true,
            Strikethrough = ChkStrikethrough.IsChecked == true,
            Charset = ResultFont.Charset
        };

        DialogResult = true;
        Close();
    }

    private void BtnCancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
