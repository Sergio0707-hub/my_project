using ClosedXML.Excel;
using Microsoft.Win32;
using System.Collections.ObjectModel;
using System.Globalization;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;

namespace Inventar;

public sealed class InventoryItem : INotifyPropertyChanged
{
    private int _fact;
    private bool _isLastScanned;

    public string Code { get; set; } = "";
    public string Name { get; set; } = "";
    public double Stock { get; set; }
    public int Fact
    {
        get => _fact;
        set
        {
            if (_fact == value) return;
            _fact = value;
            OnPropertyChanged(nameof(Fact));
            OnPropertyChanged(nameof(Diff));
            OnPropertyChanged(nameof(Status));
        }
    }
    public bool IsLastScanned
    {
        get => _isLastScanned;
        set
        {
            if (_isLastScanned == value) return;
            _isLastScanned = value;
            OnPropertyChanged(nameof(IsLastScanned));
        }
    }
    public double Diff => Fact - Stock;
    public string Status => Diff == 0 ? "Совпадает" : Diff < 0 ? "Недостача" : "Излишек";

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged(string propertyName) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
}
public sealed class SessionData
{
    public string Source { get; set; } = "";
    public List<InventoryItem> Items { get; set; } = new();
    public List<string> Undo { get; set; } = new();
}
public partial class MainWindow : Window
{
    private readonly ObservableCollection<InventoryItem> _items = new();
    private readonly Stack<string> _undo = new();
    private readonly string _sessionPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Inventar", "session.json");
    private bool _sound = true;
    private string _source = "";
    public MainWindow()
    {
        InitializeComponent();
        ItemsGrid.ItemsSource = _items;
        LoadSession();
        ScanBox.Focus();
    }
    private void Load_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog { Filter = "Книги Excel (*.xlsx)|*.xlsx", Title = "Выберите файл инвентаризации" };
        if (dlg.ShowDialog() != true) return;
        try
        {
            using var wb = new XLWorkbook(dlg.FileName);
            var dict = new Dictionary<string, InventoryItem>(StringComparer.OrdinalIgnoreCase);
            foreach (var ws in wb.Worksheets)
            {
                foreach (var row in ws.RowsUsed())
                {
                    string code = row.Cell(1).GetString().Trim();
                    if (string.IsNullOrWhiteSpace(code) || code.Equals("Артикул", StringComparison.OrdinalIgnoreCase) || code.StartsWith("Прайс-лист", StringComparison.OrdinalIgnoreCase)) continue;
                    string name = row.Cell(4).GetString().Trim();
                    double stock = 0;
                    var cell = row.Cell(14);
                    if (!cell.IsEmpty())
                    {
                        if (cell.TryGetValue<double>(out var numeric)) stock = numeric;
                        else if (!double.TryParse(cell.GetString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out stock)) continue;
                    }
                    if (!dict.ContainsKey(code)) dict[code] = new InventoryItem { Code = code, Name = name, Stock = stock };
                }
            }
            if (dict.Count == 0) { MessageBox.Show("Не найдены товары. Ожидаются артикул в A, наименование в D и остаток в N."); return; }
            if (_items.Count > 0 && MessageBox.Show("Заменить текущую инвентаризацию? Текущий прогресс будет заменён.", "Подтверждение", MessageBoxButton.YesNo) != MessageBoxResult.Yes) return;
            _items.Clear(); foreach (var item in dict.Values) _items.Add(item);
            _undo.Clear(); _source = dlg.FileName; FileLabel.Text = Path.GetFileName(_source);
            StatusText.Text = $"Загружено {_items.Count} уникальных позиций"; RefreshView(); SaveSession(); ScanBox.Focus();
        }
        catch (Exception ex) { MessageBox.Show("Ошибка чтения Excel: " + ex.Message); }
    }
    private void Scan_Click(object sender, RoutedEventArgs e) => AddScan();
    private void ScanBox_KeyDown(object sender, KeyEventArgs e) { if (e.Key == Key.Enter) { e.Handled = true; AddScan(); } }
    private void AddScan()
    {
        string code = ScanBox.Text.Trim(); ScanBox.Clear(); if (code.Length == 0) return;
        var item = _items.FirstOrDefault(x => string.Equals(x.Code, code, StringComparison.OrdinalIgnoreCase));
        if (item is null) { StatusText.Text = $"Не найден в документе: {code}"; if (_sound) System.Media.SystemSounds.Exclamation.Play(); ScanBox.Focus(); return; }
        SetLastScanned(item);
        item.Fact++; _undo.Push(item.Code); StatusText.Text = $"Добавлен: {item.Code} — {item.Name}";
        if (_sound) System.Media.SystemSounds.Asterisk.Play(); RefreshView(); SaveSession(); ScanBox.Focus();
        ItemsGrid.SelectedItem = item;
        CenterItemInView(item);
    }

    private void SetLastScanned(InventoryItem current)
    {
        foreach (var x in _items) x.IsLastScanned = false;
        current.IsLastScanned = true;
    }

    private void CenterItemInView(InventoryItem item)
    {
        // Сначала прокручиваем к элементу, чтобы WPF создал виртуализированную строку.
        ItemsGrid.ScrollIntoView(item);
        ItemsGrid.Dispatcher.BeginInvoke(new Action(() =>
        {
            ItemsGrid.UpdateLayout();
            var row = ItemsGrid.ItemContainerGenerator.ContainerFromItem(item) as System.Windows.Controls.DataGridRow;
            var viewer = FindVisualChild<System.Windows.Controls.ScrollViewer>(ItemsGrid);
            if (row is null || viewer is null) return;

            // Выравниваем строку по центру видимой области DataGrid.
            var point = row.TranslatePoint(new Point(0, 0), viewer);
            double desiredTop = Math.Max(0, (viewer.ViewportHeight - row.ActualHeight) / 2.0);
            double newOffset = viewer.VerticalOffset + point.Y - desiredTop;
            viewer.ScrollToVerticalOffset(Math.Max(0, newOffset));
            ItemsGrid.UpdateLayout();
            ItemsGrid.SelectedItem = item;
            ItemsGrid.Focus();
        }), System.Windows.Threading.DispatcherPriority.ContextIdle);
    }

    private static T? FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
    {
        int count = System.Windows.Media.VisualTreeHelper.GetChildrenCount(parent);
        for (int i = 0; i < count; i++)
        {
            var child = System.Windows.Media.VisualTreeHelper.GetChild(parent, i);
            if (child is T match) return match;
            var nested = FindVisualChild<T>(child);
            if (nested is not null) return nested;
        }
        return null;
    }

    private void Undo_Click(object sender, RoutedEventArgs e) => UndoScan();
    private void UndoScan()
    {
        if (_undo.Count == 0) { StatusText.Text = "Нет сканов для отмены"; return; }
        string code = _undo.Pop(); var item = _items.FirstOrDefault(x => x.Code == code);
        if (item is not null && item.Fact > 0)
        {
            item.Fact--;
            SetLastScanned(item);
        }
        StatusText.Text = $"Отменён последний скан: {code}"; RefreshView(); SaveSession(); ScanBox.Focus();
        if (item is not null) { ItemsGrid.SelectedItem = item; CenterItemInView(item); }
    }
    private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (Keyboard.Modifiers == ModifierKeys.Control && e.Key == Key.Z) { UndoScan(); e.Handled = true; }
    }
    private void Search_Changed(object sender, System.Windows.Controls.TextChangedEventArgs e) => RefreshView();
    private void Sound_Click(object sender, RoutedEventArgs e) { _sound = !_sound; SoundButton.Content = _sound ? "🔊 Звук: Вкл" : "🔇 Звук: Выкл"; }
    private void RefreshView()
    {
        string q = SearchBox.Text.Trim();
        ItemsGrid.Items.Refresh();
        ICollectionView view = System.Windows.Data.CollectionViewSource.GetDefaultView(_items);
        view.Filter = o => { var x=(InventoryItem)o; return q.Length==0 || x.Code.Contains(q,StringComparison.OrdinalIgnoreCase) || x.Name.Contains(q,StringComparison.OrdinalIgnoreCase); };
        TotalText.Text = _items.Count.ToString();
        MatchText.Text = _items.Count(x => x.Diff == 0).ToString();
        ShortText.Text = _items.Count(x => x.Diff < 0).ToString();
        OverText.Text = _items.Count(x => x.Diff > 0).ToString();
        ScannedText.Text = _items.Sum(x => x.Fact).ToString();
    }
    private void Finish_Click(object sender, RoutedEventArgs e)
    {
        if (_items.Count == 0) { MessageBox.Show("Сначала загрузите Excel-файл."); return; }
        int uncounted = _items.Count(x => x.Fact == 0);
        MessageBox.Show($"Всего позиций: {_items.Count}\nСовпадает: {_items.Count(x=>x.Diff==0)}\nНедостача: {_items.Count(x=>x.Diff<0)}\nИзлишек: {_items.Count(x=>x.Diff>0)}\nНе сканировались: {uncounted}", "Итоги инвентаризации");
    }
    private void Export_Click(object sender, RoutedEventArgs e)
    {
        if (_items.Count == 0) { MessageBox.Show("Нет данных для экспорта."); return; }
        var dlg = new SaveFileDialog { Filter = "CSV для Excel (*.csv)|*.csv", FileName = $"Инвентаризация_{DateTime.Now:yyyy-MM-dd_HH-mm}.csv" };
        if (dlg.ShowDialog() != true) return;
        var sb = new StringBuilder(); sb.AppendLine("Артикул;Номенклатура;Остаток по Excel;Фактически;Разница;Статус");
        foreach (var x in _items) sb.AppendLine($"{Csv(x.Code)};{Csv(x.Name)};{x.Stock.ToString(CultureInfo.InvariantCulture)};{x.Fact};{x.Diff.ToString(CultureInfo.InvariantCulture)};{Csv(x.Status)}");
        File.WriteAllText(dlg.FileName, sb.ToString(), new UTF8Encoding(true)); MessageBox.Show("Отчёт сохранён:\n" + dlg.FileName);
    }
    private static string Csv(string s) => "\"" + s.Replace("\"", "\"\"") + "\"";
    private void SaveSession()
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(_sessionPath)!);
            var data = new SessionData { Source = _source, Items = _items.ToList(), Undo = _undo.Reverse().ToList() };
            File.WriteAllText(_sessionPath, JsonSerializer.Serialize(data));
        } catch { }
    }
    private void LoadSession()
    {
        try
        {
            if (!File.Exists(_sessionPath)) return;
            var data = JsonSerializer.Deserialize<SessionData>(File.ReadAllText(_sessionPath));
            if (data?.Items.Count > 0 && MessageBox.Show($"Найдена незавершённая инвентаризация ({data.Items.Count} позиций). Продолжить?", "Восстановление", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                _source = data.Source; _items.Clear(); foreach (var item in data.Items) _items.Add(item);
                foreach (var code in data.Undo) _undo.Push(code);
                FileLabel.Text = string.IsNullOrWhiteSpace(_source) ? "Восстановленная сессия" : Path.GetFileName(_source);
                StatusText.Text = "Восстановлена последняя сессия"; RefreshView();
            }
        } catch { }
    }
    protected override void OnClosing(System.ComponentModel.CancelEventArgs e) { SaveSession(); base.OnClosing(e); }
}
