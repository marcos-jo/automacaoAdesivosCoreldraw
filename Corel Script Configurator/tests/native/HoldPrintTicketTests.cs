using System;
using FourPrintShops.PrintConfig;

internal static class HoldPrintTicketTests {
    private const string Vendor = "urn:konica-test";
    private const string Schema = "http://schemas.microsoft.com/windows/2003/08/printing/printschemaframework";
    private static void Check(bool value, string name) { if (!value) throw new Exception(name); }
    private static string Feature(string name, string option) {
        return "<f:Feature name='k:" + name + "'><f:Option name='k:" + option + "'/></f:Feature>";
    }
    private static string Ticket(string features) {
        return "<f:PrintTicket xmlns:f='" + Schema + "' xmlns:k='" + Vendor + "'>" + features + "</f:PrintTicket>";
    }
    public static int Main() {
        string normal = Feature("DocumentKOPrintMode", "Print");
        var hold = HoldTicketPolicy.Inspect(Ticket(Feature("DocumentKOOutputMethod", "Hold") + normal), Vendor);
        Check(hold.Confirmed && hold.OutputMethod == "Hold" && hold.PrintMode == "Print", "Hold with Normal");
        foreach (string option in new[] { "Print", "HoldPrint", "Unknown", "hold" }) {
            Check(!HoldTicketPolicy.Inspect(Ticket(Feature("DocumentKOOutputMethod", option) + normal), Vendor).Confirmed, "Refuse " + option);
        }
        Check(!HoldTicketPolicy.Inspect(Ticket(Feature("DocumentKOOutputMethod", "Hold") + Feature("DocumentKOPrintMode", "Wait")), Vendor).Confirmed, "Print mode not confused with output method");
        Check(!HoldTicketPolicy.Inspect(Ticket(normal), Vendor).Confirmed, "Missing output method");
        Check(!HoldTicketPolicy.Inspect(Ticket(Feature("DocumentKOOutputMethod", "Hold") + normal), "urn:wrong").Confirmed, "Unknown namespace");
        Check(!HoldTicketPolicy.Inspect(Ticket(Feature("DocumentKOOutputMethod", "Hold") + Feature("DocumentKOOutputMethod", "Print") + normal), Vendor).Confirmed, "Duplicate output methods");
        Check(!HoldTicketPolicy.Inspect("<!DOCTYPE x [<!ENTITY a SYSTEM 'file:///C:/Windows/win.ini'>]>" + Ticket(Feature("DocumentKOOutputMethod", "Hold") + normal), Vendor).Confirmed, "DTD prohibited");
        byte[] mode = new byte[240];
        using (var binding = new HoldValidation("queue", "session", 2, mode, "driver", hold)) {
            Check(binding.Matches("queue", "session", 2, mode, "driver"), "Binding matches");
            Check(!binding.Matches("queue", "session", 3, mode, "driver"), "Revision invalidates");
            Check(!binding.Matches("other", "session", 2, mode, "driver"), "Queue invalidates");
            Check(!binding.Matches("queue", "other", 2, mode, "driver"), "Session invalidates");
            Check(!binding.Matches("queue", "session", 2, mode, "new-driver"), "Driver invalidates");
            mode[230] = 17;
            Check(!binding.Matches("queue", "session", 2, mode, "driver"), "Private bytes invalidate");
            Check(!(bool)binding.Describe()["canSend"], "Retention evidence remains required");
        }
        Console.WriteLine("PASSOU: 17 verificações de Hold, campos distintos, XML, vínculo completo e bloqueio de envio.");
        return 0;
    }
}
