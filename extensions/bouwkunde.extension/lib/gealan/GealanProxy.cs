using System;
using System.IO;
using System.Reflection;
using System.Text;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.UI;

namespace GealanProxy
{
    public class ProxyApplication : IExternalApplication
    {
        private IExternalApplication _realGealan;
        
        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                string thisDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string gealanDir = Path.Combine(Path.GetDirectoryName(thisDir), "PlanersoftwareRevitPlugin");
                string gealanDll = Path.Combine(gealanDir, "PlanersoftwareRevitPlugin.dll");
                
                if (File.Exists(gealanDll))
                {
                    Assembly gealanAsm = Assembly.LoadFrom(gealanDll);
                    Type appType = gealanAsm.GetType("RevitPlugin.ExternalApplication");
                    _realGealan = (IExternalApplication)Activator.CreateInstance(appType);
                    _realGealan.OnStartup(application);
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Gealan Proxy", "Gealan plugin load warning: " + ex.Message);
            }
            
            try
            {
                RibbonPanel panel = application.CreateRibbonPanel("Gealan Tools");
                string assemblyPath = Assembly.GetExecutingAssembly().Location;
                PushButtonData btnData = new PushButtonData("ReadGealanData", "Read\nK-merken", assemblyPath, "GealanProxy.ReadGealanSchemaCommand");
                btnData.ToolTip = "Lees kozijnmerken uit Gealan ExtensibleStorage en exporteer naar JSON + mapping.txt";
                panel.AddItem(btnData);
            }
            catch (Exception) { }
            
            return Result.Succeeded;
        }
        
        public Result OnShutdown(UIControlledApplication application)
        {
            if (_realGealan != null)
                try { _realGealan.OnShutdown(application); } catch { }
            return Result.Succeeded;
        }
    }
    
    [Transaction(TransactionMode.ReadOnly)]
    public class ReadGealanSchemaCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;
            Guid schemaGuid = new Guid("36a633f0-174c-11ef-9e35-0800200c9a66");
            Schema schema = Schema.Lookup(schemaGuid);
            
            if (schema == null)
            {
                TaskDialog.Show("Gealan Reader", "Geen Gealan schema gevonden in dit document.");
                return Result.Failed;
            }
            
            IList<Element> dataElements = new FilteredElementCollector(doc)
                .WherePasses(new ExtensibleStorageFilter(schemaGuid)).ToElements();
            
            if (dataElements.Count == 0)
            {
                TaskDialog.Show("Gealan Reader", "Geen Gealan DataStorage elementen gevonden.");
                return Result.Failed;
            }
            
            IList<Field> fields = schema.ListFields();
            StringBuilder json = new StringBuilder();
            json.AppendLine("{");
            json.AppendLine("  \"schema_name\": \"" + schema.SchemaName + "\",");
            json.AppendLine("  \"element_count\": " + dataElements.Count + ",");
            json.AppendLine("  \"fields\": [");
            for (int i = 0; i < fields.Count; i++)
            {
                json.Append("    {\"name\": \"" + fields[i].FieldName + "\", \"type\": \"" + fields[i].ValueType.Name + "\"}");
                if (i < fields.Count - 1) json.Append(",");
                json.AppendLine();
            }
            json.AppendLine("  ],");
            json.AppendLine("  \"elements\": [");
            
            List<string> mappingLines = new List<string>();
            for (int e = 0; e < dataElements.Count; e++)
            {
                Element elem = dataElements[e];
                string elemName = elem.Name ?? "unknown";
                Entity entity = elem.GetEntity(schema);
                json.AppendLine("    {");
                json.AppendLine("      \"element_id\": " + elem.Id.Value + ",");
                json.AppendLine("      \"element_name\": \"" + Esc(elemName) + "\",");
                string posDesc = "";
                if (entity != null && entity.IsValid())
                {
                    json.AppendLine("      \"data\": {");
                    for (int fi = 0; fi < fields.Count; fi++)
                    {
                        Field field = fields[fi];
                        string val = ReadField(entity, field);
                        if (field.FieldName == "ProjectPositionDesc") posDesc = val;
                        json.Append("        \"" + field.FieldName + "\": \"" + Esc(val) + "\"");
                        if (fi < fields.Count - 1) json.Append(",");
                        json.AppendLine();
                    }
                    json.AppendLine("      }");
                }
                else json.AppendLine("      \"data\": null");
                json.Append("    }");
                if (e < dataElements.Count - 1) json.Append(",");
                json.AppendLine();
                if (!string.IsNullOrEmpty(posDesc))
                {
                    string kCode = posDesc.Replace("Merk ", "");
                    string kClean = kCode.Replace(" brandwerend", "");
                    bool bw = kCode.Contains("brandwerend");
                    mappingLines.Add(elemName + "|" + kClean + "|" + (bw ? "True" : "False"));
                }
            }
            json.AppendLine("  ]");
            json.AppendLine("}");
            
            string outDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "GealanReader");
            Directory.CreateDirectory(outDir);
            File.WriteAllText(Path.Combine(outDir, "gealan_schema_data.json"), json.ToString(), Encoding.UTF8);
            File.WriteAllText(Path.Combine(outDir, "mapping.txt"), string.Join(Environment.NewLine, mappingLines), Encoding.UTF8);
            
            TaskDialog.Show("Gealan K-merken Reader",
                "Export geslaagd!\n\nFamilies: " + dataElements.Count +
                "\nK-merken: " + mappingLines.Count +
                "\n\nMapping: " + Path.Combine(outDir, "mapping.txt"));
            return Result.Succeeded;
        }
        
        private string ReadField(Entity entity, Field field)
        {
            try
            {
                if (field.ValueType == typeof(string)) return entity.Get<string>(field) ?? "";
                if (field.ValueType == typeof(int)) return entity.Get<int>(field).ToString();
                if (field.ValueType == typeof(double)) return entity.Get<double>(field).ToString();
                if (field.ValueType == typeof(bool)) return entity.Get<bool>(field).ToString();
                if (field.ValueType == typeof(Guid)) return entity.Get<Guid>(field).ToString();
                if (field.ValueType == typeof(ElementId)) return entity.Get<ElementId>(field).Value.ToString();
                return "<unsupported>";
            }
            catch { return "<error>"; }
        }
        
        private string Esc(string s)
        {
            if (s == null) return "";
            return s.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\n", "\\n").Replace("\r", "\\r");
        }
    }
}
