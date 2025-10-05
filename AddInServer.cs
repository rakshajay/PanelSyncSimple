//[10/05/2025]:Raksha- Inventor Add-in (supports IPT & IAM) with IGES watcher, Export OBJ button, and clipping filter
using Inventor;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Drawing;
using System.Reflection;

namespace PanelSync.InventorAddIn
{
    [ComVisible(true)]
    [Guid("B2C7C23E-18B0-4A11-9B0B-8C6B16E30F11")]
    public class AddInServer : ApplicationAddInServer
    {
        private Inventor.Application _inv;
        private FileSystemWatcher _igesWatcher;  // [10/05/2025]:Raksha- Watch IGES folder for auto-import
        private ButtonDefinition _exportObjBtn;

        private string _hotRoot;
        private string _objDir;
        private string _igesDir;
        private string _scriptsDir;
        private string _logsDir;

        public void Activate(ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            _inv = addInSiteObject.Application;

            //[10/05/2025]:Raksha- Initialize hot-folder structure
            var desktop = System.Environment.GetFolderPath(System.Environment.SpecialFolder.UserProfile);
            _hotRoot = System.IO.Path.Combine(desktop, "OneDrive", "Desktop", "PanelSyncHot");
            _logsDir = System.IO.Path.Combine(_hotRoot, "logs");
            _objDir = System.IO.Path.Combine(_hotRoot, "Inventor", "exports", "obj");
            _igesDir = System.IO.Path.Combine(_hotRoot, "3DR", "exports", "iges");
            _scriptsDir = System.IO.Path.Combine(_hotRoot, "scripts");

            System.IO.Directory.CreateDirectory(_logsDir);
            System.IO.Directory.CreateDirectory(_objDir);
            System.IO.Directory.CreateDirectory(_igesDir);
            System.IO.Directory.CreateDirectory(_scriptsDir);

            var logPath = System.IO.Path.Combine(_logsDir, "inventor-addin.log");
            var logger = new SimpleFileLogger(logPath);

            logger.Info("Inventor Add-in activated. Watching hot-folder at " + _hotRoot);

            //[10/05/2025]:Raksha- Watch IGES folder
            _igesWatcher = new FileSystemWatcher(_igesDir, "*.igs");
            _igesWatcher.Created += (s, e) => OnNewIges(e.FullPath, logger);
            _igesWatcher.EnableRaisingEvents = true;

            //[10/05/2025]:Raksha- Seed JS scripts
            SeedScripts(logger);

            //[10/05/2025]:Raksha- Create Export OBJ button
            CreateUI();
        }

        private void SeedScripts(ILog logger)
        {
            try
            {
                // === ExportToInventor.js ===
                var exportJsPath = System.IO.Path.Combine(_scriptsDir, "ExportToInventor.js");
                if (!System.IO.File.Exists(exportJsPath))
                {
                    var exportJs = @"//[10/05/2025]:Raksha- Export visible-only geometry from OPEN 3DR project to IGES hot-folder
function log(m){try{print(m);}catch(_){}};function j(o){try{return JSON.stringify(o);}catch(_){return String(o);}}
var stamp=new Date().toISOString().replace(/[-:T.Z]/g,'');
var out='" + _igesDir.Replace("\\", "/") + @"/exp_'+stamp+'.igs';
try{
 var comps=SComp.All(SComp.VISIBLE_ONLY);
 var picked=[];
 for(var i=0;i<comps.length;i++){
   var s=comps[i].toString();
   if(((s.indexOf('Line')>=0||s.indexOf('Circle')>=0||s.indexOf('Polyline')>=0||
       s.indexOf('Multiline')>=0||s.indexOf('Plane')>=0||s.indexOf('Rectangle')>=0))
       && s.indexOf('clipping')<0){
        picked.push(comps[i]);
   }
 }
 if(picked.length===0)throw new Error('No visible geometry found.');
 var shapes=[];
 for(var k=0;k<picked.length;k++){
   var obj=picked[k];var s=obj.toString();
   if(s.indexOf('Multiline')>=0&&obj.GetNumber){
     var n=obj.GetNumber();
     for(var j=0;j<n-1;j++){
       var p1=obj.GetPoint(j);
       var p2=obj.GetPoint(j+1);
       var line=SLine.New(p1,p2);
       var conv=SCADUtil.Convert(line);
       if(conv&&conv.ErrorCode===0&&conv.Shape)shapes.push(conv.Shape);
     }
   }else{
     var conv=SCADUtil.Convert(obj);
     if(conv&&conv.ErrorCode===0&&conv.Shape)shapes.push(conv.Shape);
   }
 }
 if(shapes.length===0)throw new Error('Nothing convertible to IGES.');
 var mtx=SMatrix.New();mtx.InitIdentity();
 log('Exporting IGES → '+out);
 var rc=SCADUtil.Export(out,shapes,mtx);
 if(!rc||rc.ErrorCode!==0)throw new Error('Export failed: '+j(rc));
 log('✅ IGES exported: '+out);
}catch(err){log('ERROR: '+err);throw err;}";
                    System.IO.File.WriteAllText(exportJsPath, exportJs);
                    logger.Info("Seeded ExportToInventor.js (with clipping filter) into scripts folder");
                }

                // === latest_obj.js ===
                var objJsPath = System.IO.Path.Combine(_scriptsDir, "latest_obj.js");
                if (!System.IO.File.Exists(objJsPath))
                {
                    var objJs = @"//[10/05/2025]:Raksha- Import latest.obj into OPEN 3DR project
var objPath='" + System.IO.Path.Combine(_objDir, "latest.obj").Replace("\\", "/") + @"';
function log(m){try{print(m);}catch(_){}};if(!objPath){throw'objPath is missing!';}
log('Importing OBJ: '+objPath);
var rc=SPoly.FromFile(objPath);
if(!rc||rc.ErrorCode!==0){throw'SPoly.FromFile failed: '+JSON.stringify(rc);}
if(!rc.PolyTbl||rc.PolyTbl.length===0){throw'No meshes found in OBJ';}
for(var i=0;i<rc.PolyTbl.length;i++){var mesh=rc.PolyTbl[i];mesh.AddToDoc();log('Added mesh: '+mesh.GetName()+' ('+i+')');}
try{var vs=SViewSet.New(true);vs.Update(true);}catch(_){}
log('✅ Imported OBJ and updated view.');";
                    System.IO.File.WriteAllText(objJsPath, objJs);
                    logger.Info("Seeded latest_obj.js into scripts folder");
                }
            }
            catch (Exception ex)
            {
                logger.Error("Failed seeding scripts", ex);
            }
        }

        //[10/05/2025]:Raksha- IGES auto-import handler (same as before)
        private void OnNewIges(string igesPath, ILog logger)
        {
            try
            {
                logger.Info("Detected new IGES file: " + igesPath);
                System.Threading.Thread.Sleep(1000);

                Document activeDoc = _inv.ActiveDocument;
                if (activeDoc == null)
                {
                    var iptPath = System.IO.Path.Combine(_hotRoot, "Inventor", "projects", "latest.ipt");
                    activeDoc = _inv.Documents.Add(DocumentTypeEnum.kPartDocumentObject,
                        _inv.FileManager.GetTemplateFile(DocumentTypeEnum.kPartDocumentObject));
                    activeDoc.SaveAs(iptPath, false);
                }

                if (activeDoc is PartDocument partDoc)
                {
                    var compDef = partDoc.ComponentDefinition;
                    var importedDef = compDef.ReferenceComponents.ImportedComponents.CreateDefinition(igesPath);
                    partDoc.UnitsOfMeasure.LengthUnits = UnitsTypeEnum.kMillimeterLengthUnits;
                    compDef.ReferenceComponents.ImportedComponents.Add(importedDef);
                    partDoc.Save();
                    logger.Info("✅ IGES imported into Part: " + System.IO.Path.GetFileName(igesPath));
                }
                else if (activeDoc is AssemblyDocument asmDoc)
                {
                    var iptPath = System.IO.Path.Combine(_hotRoot, "Inventor", "projects", "autoimport.ipt");
                    PartDocument newPart = (PartDocument)_inv.Documents.Add(
                        DocumentTypeEnum.kPartDocumentObject,
                        _inv.FileManager.GetTemplateFile(DocumentTypeEnum.kPartDocumentObject));
                    newPart.SaveAs(iptPath, false);

                    var compDef = newPart.ComponentDefinition;
                    var def = compDef.ReferenceComponents.ImportedComponents.CreateDefinition(igesPath);
                    compDef.ReferenceComponents.ImportedComponents.Add(def);
                    newPart.Save();

                    var pos = _inv.TransientGeometry.CreateMatrix();
                    asmDoc.ComponentDefinition.Occurrences.Add(iptPath, pos);
                    asmDoc.Save();

                    logger.Info("✅ IGES imported into Assembly as new part: " + System.IO.Path.GetFileName(igesPath));
                }
            }
            catch (Exception ex)
            {
                logger.Error("Error during IGES auto-import", ex);
            }
        }

        //[10/05/2025]:Raksha- Create Export OBJ button (added to both Part & Assembly ribbons)
        private void CreateUI()
        {
            var cmdMgr = _inv.CommandManager;
            var controlDefs = cmdMgr.ControlDefinitions;

            var asm = Assembly.GetExecutingAssembly();
            using (var bmp = new Bitmap(asm.GetManifestResourceStream("PanelSync.InventorAddIn.Resources.OBJIcon.png")))
            {
                var icon16 = PictureDispConverter.ToIPictureDisp(new Bitmap(bmp, new Size(16, 16)));
                var icon32 = PictureDispConverter.ToIPictureDisp(new Bitmap(bmp, new Size(32, 32)));

                _exportObjBtn = controlDefs.AddButtonDefinition(
                    "Export OBJ",
                    "PanelSyncExportObjBtn",
                    CommandTypesEnum.kShapeEditCmdType,
                    Guid.NewGuid().ToString(),
                    "Export active document as OBJ to PanelSyncHot",
                    "Exports active part or assembly to OBJ",
                    icon16,
                    icon32
                );

                _exportObjBtn.OnExecute += ExportObjBtn_OnExecute;

                //[10/05/2025]:Raksha- Add to Part ribbon
                var partRibbon = _inv.UserInterfaceManager.Ribbons["Part"];
                var partTools = partRibbon.RibbonTabs["id_TabTools"];
                var partPanel = partTools.RibbonPanels.Add("PanelSync", "PanelSyncPartPanel", Guid.NewGuid().ToString(), "", false);
                partPanel.CommandControls.AddButton(_exportObjBtn, true);

                //[10/05/2025]:Raksha- Add to Assembly ribbon
                var asmRibbon = _inv.UserInterfaceManager.Ribbons["Assembly"];
                var asmTools = asmRibbon.RibbonTabs["id_TabTools"];
                var asmPanel = asmTools.RibbonPanels.Add("PanelSync", "PanelSyncAsmPanel", Guid.NewGuid().ToString(), "", false);
                asmPanel.CommandControls.AddButton(_exportObjBtn, true);
            }
        }

        //[10/05/2025]:Raksha- Unified OBJ export handler for IPT & IAM
        private void ExportObjBtn_OnExecute(NameValueMap Context)
        {
            try
            {
                Document doc = _inv.ActiveDocument;
                if (doc == null)
                {
                    _inv.StatusBarText = "No document open.";
                    return;
                }

                var objPath = System.IO.Path.Combine(_objDir, "latest.obj");
                var ctx = _inv.TransientObjects.CreateTranslationContext();
                ctx.Type = IOMechanismEnum.kFileBrowseIOMechanism;
                var data = _inv.TransientObjects.CreateDataMedium();
                data.FileName = objPath;
                var options = _inv.TransientObjects.CreateNameValueMap();
                var addin = _inv.ApplicationAddIns.ItemById["{F539FB09-FC01-4260-A429-1818B14D6BAC}"]; // OBJ Translator
                var trans = (TranslatorAddIn)addin;

                if (trans.HasSaveCopyAsOptions[doc, ctx, options])
                {
                    SetOpt(options, "ExportAllSolids", true);
                    SetOpt(options, "ExportSelection", 0);
                    SetOpt(options, "Resolution", 5);
                    SetOpt(options, "SurfaceType", 0);
                    trans.SaveCopyAs(doc, ctx, options, data);
                }

                System.IO.File.SetLastWriteTimeUtc(objPath, DateTime.UtcNow);
                _inv.StatusBarText = "✅ OBJ exported: " + objPath;
            }
            catch (Exception ex)
            {
                _inv.StatusBarText = "OBJ export failed: " + ex.Message;
            }
        }

        private void SetOpt(NameValueMap opts, string key, object value)
        {
            try { opts.Value[key] = value; }
            catch { try { opts.Add(key, value); } catch { } }
        }

        public void Deactivate()
        {
            try { _igesWatcher?.Dispose(); } catch { }
            _exportObjBtn?.Delete();
            _inv = null; _igesWatcher = null;
        }

        public void ExecuteCommand(int commandID) { }
        public object Automation => null;

        //[10/05/2025]:Raksha- Helper to convert .NET Bitmap → Inventor COM icon
        public class PictureDispConverter : AxHost
        {
            private PictureDispConverter() : base("") { }
            public static stdole.IPictureDisp ToIPictureDisp(System.Drawing.Image image)
            {
                return (stdole.IPictureDisp)AxHost.GetIPictureDispFromPicture(image);
            }
        }
    }
}
