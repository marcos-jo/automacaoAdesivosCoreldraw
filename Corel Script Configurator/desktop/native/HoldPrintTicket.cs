using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Printing;
using System.Printing.Interop;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Win32;

namespace FourPrintShops.PrintConfig {
    internal sealed class HoldDecision {
        internal string OutputMethod;
        internal string PrintMode;
        internal string Reason;
        internal bool Confirmed;
        internal bool Conclusive;
    }

    internal static class HoldTicketPolicy {
        private const string Schema = "http://schemas.microsoft.com/windows/2003/08/printing/printschemaframework";

        internal static HoldDecision Inspect(string xml, string expectedNamespace) {
            var decision = new HoldDecision();
            try {
                var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 4 * 1024 * 1024 };
                XDocument document;
                using (var text = new StringReader(xml))
                using (var reader = XmlReader.Create(text, settings)) {
                    document = XDocument.Load(reader);
                }
                XNamespace schema = Schema;
                if (document.Root == null || document.Root.Name != schema + "PrintTicket") {
                    throw new InvalidOperationException();
                }
                decision.OutputMethod = ReadOption(document.Root, "DocumentKOOutputMethod", expectedNamespace);
                decision.PrintMode = ReadOption(document.Root, "DocumentKOPrintMode", expectedNamespace);
                decision.Conclusive = new[] { "Print", "Hold", "HoldPrint" }.Contains(decision.OutputMethod)
                    && new[] { "Print", "ProofMode", "Proof1stSheet", "Wait", "Spool" }.Contains(decision.PrintMode);
                decision.Confirmed = decision.OutputMethod == "Hold" && decision.PrintMode == "Print";
                decision.Reason = decision.Confirmed
                    ? "Hold confirmado no DEVMODE desta sessão. A retenção na saída gerada ainda aguarda comprovação."
                    : "Envio bloqueado. É necessário Output Method = Hold e Print Mode = Normal.";
            } catch {
                decision.Confirmed = false;
                decision.Reason = "Envio bloqueado. O PrintTicket não permitiu comprovar inequivocamente o modo Hold.";
            }
            return decision;
        }

        private static string ReadOption(XElement root, string featureName, string expectedNamespace) {
            XNamespace schema = Schema;
            // Somente Feature no nível do trabalho convertido. Duplicatas/escopos ambíguos são recusados.
            var matching = root.Descendants(schema + "Feature")
                .Where(element => LocalName(element.Attribute("name")) == featureName).ToArray();
            if (matching.Length != 1 || matching[0].Parent != root ||
                Namespace(matching[0], matching[0].Attribute("name")) != expectedNamespace) {
                throw new InvalidOperationException();
            }
            var options = matching[0].Elements(schema + "Option").ToArray();
            if (options.Length != 1 || Namespace(options[0], options[0].Attribute("name")) != expectedNamespace) {
                throw new InvalidOperationException();
            }
            return LocalName(options[0].Attribute("name"));
        }

        private static string LocalName(XAttribute attribute) {
            if (attribute == null) throw new InvalidOperationException();
            var parts = attribute.Value.Split(':');
            if (parts.Length != 2 || parts.Any(String.IsNullOrEmpty)) throw new InvalidOperationException();
            return parts[1];
        }

        private static string Namespace(XElement element, XAttribute attribute) {
            LocalName(attribute);
            var value = element.GetNamespaceOfPrefix(attribute.Value.Split(':')[0]);
            if (value == null) throw new InvalidOperationException();
            return value.NamespaceName;
        }
    }

    internal sealed class KonicaDriverIdentity {
        internal string Fingerprint;
        internal string TicketNamespace;
        internal const string Queue = "KONICA MINOLTA C6100PS";
        internal const string Driver = "KONICA MINOLTA C6100/C6085PS";

        internal static KonicaDriverIdentity Read(IntPtr printer, string queue) {
            if (queue != Queue) throw new InvalidOperationException("A verificação de Hold requer a fila Konica de referência.");
            uint needed;
            Native.GetPrinterW(printer, 2, IntPtr.Zero, 0, out needed);
            int capacity = checked((int)needed);
            IntPtr buffer = Native.Allocate(capacity);
            string port, processor, datatype;
            try {
                if (!Native.GetPrinterW(printer, 2, buffer, (uint)capacity, out needed)) throw new InvalidOperationException();
                var info = (Native.Printer2)Marshal.PtrToStructure(buffer, typeof(Native.Printer2));
                if (Native.Text(info.Driver) != Driver || Native.Text(info.Name) != queue) throw new InvalidOperationException();
                port = Native.Text(info.Port);
                processor = Native.Text(info.Processor);
                datatype = Native.Text(info.Datatype);
            } finally {
                Native.Free(buffer, capacity);
            }
            string registryPath = @"SYSTEM\CurrentControlSet\Control\Print\Environments\Windows x64\Drivers\Version-3\" + Driver;
            using (var key = Registry.LocalMachine.OpenSubKey(registryPath, false)) {
                if (key == null) throw new InvalidOperationException();
                string core = ResolveDriverFile(Convert.ToString(key.GetValue("Driver")));
                string ui = ResolveDriverFile(Convert.ToString(key.GetValue("Configuration File")));
                string ppd = ResolveDriverFile(Convert.ToString(key.GetValue("Data File")));
                string oem = ResolveDriverFile("KOI6100_ui.dll");
                if (FileVersionInfo.GetVersionInfo(oem).FileVersion != "5.0.504") throw new InvalidOperationException();
                var files = new List<string> { core, ui, ppd, oem };
                var dependencies = key.GetValue("Dependent Files") as string[];
                if (dependencies != null) files.AddRange(dependencies.Select(ResolveDriverFile));
                var identityText = new StringBuilder(queue + "\n" + Driver + "\n" + port + "\n" + processor + "\n" + datatype + "\n");
                foreach (string file in files.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(value => value, StringComparer.OrdinalIgnoreCase)) {
                    if (new FileInfo(file).Length > 32 * 1024 * 1024) throw new InvalidOperationException();
                    using (var input = File.OpenRead(file))
                    using (var hash = SHA256.Create()) {
                        identityText.Append(file).Append(':').Append(Convert.ToBase64String(hash.ComputeHash(input))).Append('\n');
                    }
                }
                var version = FileVersionInfo.GetVersionInfo(core);
                // O provedor usa a versão textual do PScript. Os componentes fixos desse DLL são diferentes.
                var versionMatch = Regex.Match(version.FileVersion ?? "", @"^(\d{1,5}\.\d{1,5}\.\d{1,5}\.\d{1,5})(?:\s|$)");
                if (!versionMatch.Success) throw new InvalidOperationException();
                string coreVersion = versionMatch.Groups[1].Value;
                using (var hash = SHA256.Create()) {
                    return new KonicaDriverIdentity {
                        Fingerprint = Convert.ToBase64String(hash.ComputeHash(Encoding.UTF8.GetBytes(identityText.ToString()))),
                        TicketNamespace = "http://schemas.microsoft.com/windows/printing/oemdriverpt/KONICA_MINOLTA C6100/C6085PS/" + coreVersion + "/"
                    };
                }
            }
        }

        private static string ResolveDriverFile(string value) {
            if (String.IsNullOrWhiteSpace(value)) throw new InvalidOperationException();
            string windows = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
            string path = Environment.ExpandEnvironmentVariables(value);
            if (!Path.IsPathRooted(path)) path = Path.Combine(windows, @"System32\spool\drivers\x64\3", path);
            path = Path.GetFullPath(path);
            if (!path.StartsWith(windows + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException();
            return path;
        }
    }

    internal sealed class HoldValidation : IDisposable {
        private byte[] modeHash;
        private readonly string queue, session, driver;
        private readonly int revision;
        private readonly HoldDecision decision;

        internal HoldValidation(string queue, string session, int revision, byte[] mode, string driver, HoldDecision decision) {
            this.queue = queue;
            this.session = session;
            this.revision = revision;
            this.driver = driver;
            this.decision = decision;
            using (var hash = SHA256.Create()) modeHash = hash.ComputeHash(mode);
        }

        internal bool Matches(string queue, string session, int revision, byte[] mode, string driver) {
            if (modeHash == null || this.queue != queue || this.session != session || this.revision != revision || this.driver != driver) return false;
            byte[] candidate;
            using (var hash = SHA256.Create()) candidate = hash.ComputeHash(mode);
            try { return candidate.SequenceEqual(modeHash); }
            finally { Array.Clear(candidate, 0, candidate.Length); }
        }

        internal Dictionary<string, object> Describe() {
            return new Dictionary<string, object> {
                {"verified", decision.Conclusive}, {"holdConfirmed", decision.Confirmed}, {"outputMethod", decision.OutputMethod},
                {"printMode", decision.PrintMode}, {"reason", decision.Reason}, {"outputValidated", false}, {"canSend", false},
                {"session", session}, {"revision", revision}, {"printer", queue}
            };
        }

        internal static Dictionary<string, object> Unverified(string reason) {
            return new Dictionary<string, object> {
                {"verified", false}, {"holdConfirmed", false}, {"outputValidated", false}, {"canSend", false}, {"reason", reason}
            };
        }

        public void Dispose() {
            if (modeHash != null) Array.Clear(modeHash, 0, modeHash.Length);
            modeHash = null;
        }
    }

    internal static class HoldTicketReader {
        internal static HoldDecision Read(string queue, byte[] mode, string ticketNamespace) {
            using (var converter = new PrintTicketConverter(queue, 1)) {
                var ticket = converter.ConvertDevModeToPrintTicket(mode, PrintTicketScope.JobScope);
                using (var stream = ticket.GetXmlStream())
                using (var reader = new StreamReader(stream)) {
                    return HoldTicketPolicy.Inspect(reader.ReadToEnd(), ticketNamespace);
                }
            }
        }
    }
}
