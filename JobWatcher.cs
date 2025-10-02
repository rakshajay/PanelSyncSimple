//[10/01/2025]:Raksha- Simplified JobWatcher (watches IGES folder + Jobs JSON)
using Inventor;
using Newtonsoft.Json;
using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using IOFile = System.IO.File;
using IOPath = System.IO.Path;
using SysEnv = System.Environment;

namespace PanelSync.InventorAddIn
{
    internal class JobWatcher : IDisposable
    {
        private readonly Application _inv;
        private readonly ILog _log;
        private readonly FileSystemWatcher _jobWatcher;
        private readonly FileSystemWatcher _igesWatcher;

        private readonly string _hotRoot;
        private readonly string _jobsDir;
        private readonly string _igesDir;
        private readonly string _objDir;
        private readonly string _projDir;

        public JobWatcher(Application inv, ILog log)
        {
            _inv = inv;
            _log = log;

            // 🔄 Ensure hot-folder structure
            var desktop = SysEnv.GetFolderPath(SysEnv.SpecialFolder.UserProfile);
            _hotRoot = IOPath.Combine(desktop, "OneDrive", "Desktop", "PanelSyncHot");
            _jobsDir = IOPath.Combine(_hotRoot, "Jobs");
            _igesDir = IOPath.Combine(_hotRoot, "3DR", "exports", "iges");
            _objDir = IOPath.Combine(_hotRoot, "Inventor", "exports", "obj");
            _projDir = IOPath.Combine(_hotRoot, "Inventor", "Projects");

            Directory.CreateDirectory(_jobsDir);
            Directory.CreateDirectory(_igesDir);
            Directory.CreateDirectory(_objDir);
            Directory.CreateDirectory(_projDir);

            _log.Info("Hot-folder initialized at: " + _hotRoot);

            // === JSON Jobs Watcher (OBJ export still needs it) ===
            _jobWatcher = new FileSystemWatcher(_jobsDir, "*.json");
            _jobWatcher.IncludeSubdirectories = false;
            _jobWatcher.NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite;
            _jobWatcher.Created += OnCreated;
            _jobWatcher.Changed += OnCreated;
            _jobWatcher.Renamed += OnCreated;
            _jobWatcher.EnableRaisingEvents = true;

            _log.Info("Watching Jobs folder: " + _jobsDir);

            // === IGES Watcher (direct import, no JSON needed) ===
            _igesWatcher = new FileSystemWatcher(_igesDir, "*.igs");
            _igesWatcher.IncludeSubdirectories = false;
            _igesWatcher.Created += async (s, e) => await OnNewIgesAsync(e.FullPath);
            _igesWatcher.EnableRaisingEvents = true;

            _log.Info("Watching IGES folder: " + _igesDir);
        }

        // === Process JSON Job ===
        private void OnCreated(object sender, FileSystemEventArgs e)
        {
            ThreadPool.QueueUserWorkItem(_ => ProcessJobFile(e.FullPath));
        }

        private void ProcessJobFile(string path)
        {
            try
            {
                string json = IOFile.ReadAllText(path);
                var kindOnly = JsonConvert.DeserializeObject<dynamic>(json);
                string kind = kindOnly?.Kind;

                if (kind == "ExportPanelAsOBJ")
                {
                    var job = JsonConvert.DeserializeObject<ExportPanelAsObjJob>(json);
                    if (job != null && job.IsValid)
                        ExecuteExportPanelAsObj(job);
                }
                else
                {
                    _log.Warn("Unknown or unsupported job kind: " + kind);
                }
            }
            catch (Exception ex)
            {
                _log.Error("Error processing job " + path, ex);
            }
        }

        // === IGES direct import ===
        private async Task OnNewIgesAsync(string igesPath)
        {
            try
            {
                _log.Info("Detected new IGES file: " + igesPath);

                bool ok = await FileStability.WaitUntilStableAsync(igesPath, 500, 150, 10000);
                if (!ok)
                {
                    _log.Warn("Skipped IGES (not stable): " + igesPath);
                    return;
                }

                // Use active part doc if available, else create new
                PartDocument doc = null;
                if (_inv.ActiveDocument is PartDocument activePart)
                {
                    doc = activePart;
                }
                else
                {
                    var iptPath = IOPath.Combine(_projDir, "latest.ipt");
                    doc = (PartDocument)_inv.Documents.Add(
                        DocumentTypeEnum.kPartDocumentObject,
                        _inv.FileManager.GetTemplateFile(DocumentTypeEnum.kPartDocumentObject));
                    doc.SaveAs(iptPath, false);
                }

                var compDef = doc.ComponentDefinition;
                var importedDef = compDef.ReferenceComponents.ImportedComponents.CreateDefinition(igesPath);

                doc.UnitsOfMeasure.LengthUnits = UnitsTypeEnum.kMillimeterLengthUnits;
                compDef.ReferenceComponents.ImportedComponents.Add(importedDef);
                doc.Save();

                _log.Info($"✅ IGES auto-import complete: {IOPath.GetFileName(igesPath)}");
            }
            catch (Exception ex)
            {
                _log.Error("Error during IGES auto-import", ex);
            }
        }

        // === OBJ Export (via job JSON) ===
        private void ExecuteExportPanelAsObj(ExportPanelAsObjJob job)
        {
            try
            {
                PartDocument doc = null;
                foreach (Document d in _inv.Documents)
                {
                    if (string.Equals(d.FullFileName, job.IptPath, StringComparison.OrdinalIgnoreCase))
                    {
                        doc = (PartDocument)d;
                        break;
                    }
                }
                if (doc == null)
                {
                    _log.Warn("OBJ export skipped: " + job.IptPath + " is not open in Inventor.");
                    return;
                }

                var compDef = doc.ComponentDefinition as PartComponentDefinition;
                if (compDef == null || compDef.SurfaceBodies == null || compDef.SurfaceBodies.Count == 0)
                {
                    _log.Warn("OBJ export skipped: no solid bodies found.");
                    return;
                }

                var ctx = _inv.TransientObjects.CreateTranslationContext();
                ctx.Type = IOMechanismEnum.kFileBrowseIOMechanism;

                var data = _inv.TransientObjects.CreateDataMedium();
                Directory.CreateDirectory(job.OutFolder);

                var baseName = IOPath.GetFileNameWithoutExtension(job.IptPath);
                var objName = $"{baseName}_{job.PanelId}_r{job.Rev}.obj";
                var objPath = IOPath.Combine(job.OutFolder, objName);

                if (IOFile.Exists(objPath))
                {
                    try { IOFile.Delete(objPath); } catch { }
                }

                data.FileName = objPath;

                var options = _inv.TransientObjects.CreateNameValueMap();
                var addin = _inv.ApplicationAddIns.ItemById["{F539FB09-FC01-4260-A429-1818B14D6BAC}"];
                var trans = (TranslatorAddIn)addin;

                int solidCount = doc.ComponentDefinition.SurfaceBodies
                   .OfType<SurfaceBody>()
                   .Count(b => b.IsSolid);

                if (solidCount == 0)
                {
                    _log.Warn("⚠️ No solid bodies found to export as OBJ.");
                    return;
                }

                if (trans.HasSaveCopyAsOptions[doc, ctx, options])
                {
                    SetOpt(options, "ExportAllSolids", true);
                    SetOpt(options, "ExportSelection", 0);
                    SetOpt(options, "Resolution", 5);
                    SetOpt(options, "SurfaceType", 0);

                    trans.SaveCopyAs(doc, ctx, options, data);
                }

                IOFile.SetLastWriteTimeUtc(objPath, DateTime.UtcNow);
                _log.Info("Exported OBJ -> " + objPath);
            }
            catch (Exception ex)
            {
                _log.Error("ExportPanelAsOBJ failed", ex);
            }
        }

        private void SetOpt(NameValueMap opts, string key, object value)
        {
            try { opts.Value[key] = value; }
            catch { try { opts.Add(key, value); } catch { } }
        }

        public void Dispose()
        {
            _jobWatcher?.Dispose();
            _igesWatcher?.Dispose();
        }
    }

    // === DTOs ===
    internal class ExportPanelAsObjJob
    {
        public string Kind { get; set; } = "ExportPanelAsOBJ";
        public string IptPath { get; set; }
        public string OutFolder { get; set; }
        public string PanelId { get; set; } = "P001";
        public string Rev { get; set; } = "A";
        public bool BringToFront { get; set; } = true;

        public bool IsValid =>
            Kind == "ExportPanelAsOBJ"
            && !string.IsNullOrWhiteSpace(IptPath)
            && !string.IsNullOrWhiteSpace(OutFolder);
    }

    // Simple stability check
    internal static class FileStability
    {
        public static async Task<bool> WaitUntilStableAsync(string path, int firstDelayMs, int pollMs, int timeoutMs)
        {
            var start = DateTime.UtcNow;
            await Task.Delay(firstDelayMs);
            long lastLen = -1; DateTime lastWrite = DateTime.MinValue; int stableCount = 0;
            while (true)
            {
                if (!IOFile.Exists(path)) return false;
                var fi = new FileInfo(path);
                if (fi.Length == lastLen && fi.LastWriteTimeUtc == lastWrite) stableCount++;
                else { stableCount = 0; lastLen = fi.Length; lastWrite = fi.LastWriteTimeUtc; }
                if (stableCount >= 2) return true;
                if ((DateTime.UtcNow - start).TotalMilliseconds > timeoutMs) return false;
                await Task.Delay(pollMs);
            }
        }
    }
}
