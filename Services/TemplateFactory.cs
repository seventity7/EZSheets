using EZSheets.Models;

namespace EZSheets.Services;

public static class TemplateFactory
{
    public static (int Rows, int Columns, string DefaultTitle) GetTemplateDefaults(SheetTemplateKind kind)
        => kind switch
        {
            SheetTemplateKind.Schedule => (30, 8, "Schedule"),
            SheetTemplateKind.Staff => (40, 8, "Staff"),
            SheetTemplateKind.Bartender => (40, 10, "Bartender"),
            SheetTemplateKind.Waitlist => (50, 7, "Waitlist"),
            SheetTemplateKind.Blackjack => (35, 10, "Blackjack"),
            SheetTemplateKind.VenueAddressBook => (35, 8, "Venue Address Book"),
            SheetTemplateKind.Finance => (35, 8, "Finance"),
            _ => (30, 12, "New Sheet"),
        };

    public static AdvancedSheetEnvelope CreateTemplate(SheetTemplateKind kind)
    {
        var envelope = new AdvancedSheetEnvelope();
        envelope.Settings.TemplateKind = kind;

        switch (kind)
        {
            case SheetTemplateKind.Schedule:
                ApplyScheduleTemplate(envelope);
                break;
            case SheetTemplateKind.Staff:
                ApplyStaffTemplate(envelope);
                break;
            case SheetTemplateKind.Bartender:
                ApplyBartenderTemplate(envelope);
                break;
            case SheetTemplateKind.Waitlist:
                ApplyWaitlistTemplate(envelope);
                break;
            case SheetTemplateKind.Blackjack:
                ApplyBlackjackTemplate(envelope);
                break;
            case SheetTemplateKind.VenueAddressBook:
                ApplyVenueAddressTemplate(envelope);
                break;
            case SheetTemplateKind.Finance:
                ApplyFinanceTemplate(envelope);
                break;
            default:
                ApplyBlankTemplate(envelope);
                break;
        }

        return envelope;
    }

    private static void ApplyBlankTemplate(AdvancedSheetEnvelope envelope)
    {
        envelope.Document = CreateDocument(30, 12, "Sheet 1", "Sheet 2");
    }

    private static void ApplyScheduleTemplate(AdvancedSheetEnvelope envelope)
    {
        envelope.Document = CreateDocument(30, 8, "Schedule", "Notes");
        var tab = envelope.Document.Tabs[0];
        SetHeaderRow(tab, new[] { "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday" }, startColumn: 1);
        StyleHeaderRange(tab, 1, 1, 1, 7, 0xFFFFFFFF, 0xFF3E3E3E);
    }

    private static void ApplyStaffTemplate(AdvancedSheetEnvelope envelope)
    {
        envelope.Document = CreateDocument(40, 8, "Staff", "Contacts");
        var tab = envelope.Document.Tabs[0];
        SetHeaderRow(tab, new[] { "Name", "Role", "Shift Start", "Shift End", "World", "Status", "Notes" }, startColumn: 1);
        StyleHeaderRange(tab, 1, 1, 1, 7, 0xFFFFFFFF, 0xFF365A8C);
    }

    private static void ApplyBartenderTemplate(AdvancedSheetEnvelope envelope)
    {
        envelope.Document = CreateDocument(40, 10, "Bartender", "Stock", "Events");
        var tab = envelope.Document.Tabs[0];
        SetHeaderRow(tab, new[] { "Customer", "Drink", "Order Time", "Served", "Tips", "Payment", "Notes", "Status", "Bartender" }, startColumn: 1);
        StyleHeaderRange(tab, 1, 1, 1, 9, 0xFFFFFFFF, 0xFF6A3D1B);
        envelope.Settings.Category = "Venue";
    }

    private static void ApplyWaitlistTemplate(AdvancedSheetEnvelope envelope)
    {
        envelope.Document = CreateDocument(50, 7, "Waitlist", "Archive");
        var tab = envelope.Document.Tabs[0];
        SetHeaderRow(tab, new[] { "Position", "Name", "Party Size", "Requested At", "Status", "Notes" }, startColumn: 1);
        StyleHeaderRange(tab, 1, 1, 1, 6, 0xFFFFFFFF, 0xFF5C2D91);
    }

    private static void ApplyBlackjackTemplate(AdvancedSheetEnvelope envelope)
    {
        envelope.Document = CreateDocument(35, 10, "Blackjack", "Summary");
        var tab = envelope.Document.Tabs[0];
        SetHeaderRow(tab, new[] { "Dealer", "Player", "Bet", "Result", "Payout", "Tips", "Session", "Notes", "Timestamp" }, startColumn: 1);
        StyleHeaderRange(tab, 1, 1, 1, 9, 0xFFFFFFFF, 0xFF1A6B2E);
    }

    private static void ApplyVenueAddressTemplate(AdvancedSheetEnvelope envelope)
    {
        envelope.Document = CreateDocument(35, 8, "Address", "Travel Notes");
        var tab = envelope.Document.Tabs[0];
        SetHeaderRow(tab, new[] { "House", "Data Center", "World", "Residential", "Ward", "Plot", "Notes" }, startColumn: 1);
        StyleHeaderRange(tab, 1, 1, 1, 7, 0xFFFFFFFF, 0xFF335F86);
    }

    private static void ApplyFinanceTemplate(AdvancedSheetEnvelope envelope)
    {
        envelope.Document = CreateDocument(35, 8, "Finance", "Summary");
        var tab = envelope.Document.Tabs[0];
        SetHeaderRow(tab, new[] { "Date", "Description", "Debit", "Credit", "Balance", "Category", "Notes" }, startColumn: 1);
        StyleHeaderRange(tab, 1, 1, 1, 7, 0xFFFFFFFF, 0xFF4B6E1C);
    }

    private static SheetDocument CreateDocument(int rows, int cols, params string[] tabNames)
    {
        var document = new SheetDocument();
        foreach (var tabName in tabNames)
        {
            document.Tabs.Add(new SheetTabData { Name = tabName });
        }

        if (document.Tabs.Count == 0)
        {
            document.Tabs.Add(new SheetTabData { Name = "Sheet 1" });
        }

        document.ActiveTabIndex = 0;
        return document;
    }

    private static void SetHeaderRow(SheetTabData tab, IReadOnlyList<string> headers, int startColumn)
    {
        for (var i = 0; i < headers.Count; i++)
        {
            var key = ToCellKey(1, startColumn + i);
            tab.GetOrCreateCell(key).Value = headers[i];
        }
    }

    private static void StyleHeaderRange(SheetTabData tab, int startRow, int startColumn, int endRow, int endColumn, uint textColor, uint backgroundColor)
    {
        for (var row = startRow; row <= endRow; row++)
        {
            for (var col = startColumn; col <= endColumn; col++)
            {
                var cell = tab.GetOrCreateCell(ToCellKey(row, col));
                cell.TextColor = textColor;
                cell.BackgroundColor = backgroundColor;
                cell.Bold = true;
            }
        }
    }

    private static string ToCellKey(int row, int col)
        => $"R{row}C{col}";

    private static string ToColumnLabel(int column)
    {
        var value = column;
        var label = string.Empty;
        while (value > 0)
        {
            value--;
            label = (char)('A' + (value % 26)) + label;
            value /= 26;
        }

        return label;
    }
}
