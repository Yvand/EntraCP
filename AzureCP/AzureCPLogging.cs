using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using System;
using System.Collections.Generic;
using System.Linq;

namespace azurecp
{
    /// <summary>
    /// Implemented as documented in http://www.sbrickey.com/Tech/Blog/Post/Custom_Logging_in_SharePoint_2010
    /// </summary>
    [System.Runtime.InteropServices.GuidAttribute("3DD2C709-C860-4A20-8AF2-0FDDAA9C406B")]
    public class AzureCPLogging : SPDiagnosticsServiceBase
    {
        public static string DiagnosticsAreaName = "AzureCP";

        public enum Categories
        {
            [CategoryName("Core"),
            DefaultTraceSeverity(TraceSeverity.Medium),
                //DefaultTraceSeverity(TraceSeverity.VerboseEx),
            DefaultEventSeverity(EventSeverity.Error)]
            Core,
            [CategoryName("Configuration"),
            DefaultTraceSeverity(TraceSeverity.Medium),
            DefaultEventSeverity(EventSeverity.Error)]
            Configuration,
            [CategoryName("Lookup"),
             DefaultTraceSeverity(TraceSeverity.Medium),
             DefaultEventSeverity(EventSeverity.Error)]
            Lookup,
            [CategoryName("Claims Picking"),
             DefaultTraceSeverity(TraceSeverity.Medium),
             DefaultEventSeverity(EventSeverity.Error)]
            Claims_Picking,
            [CategoryName("Claims Augmentation"),
             DefaultTraceSeverity(TraceSeverity.Medium),
             DefaultEventSeverity(EventSeverity.Error)]
            Claims_Augmentation,
            [CategoryName("Rehydration"),
             DefaultTraceSeverity(TraceSeverity.Medium),
             DefaultEventSeverity(EventSeverity.Error)]
            Rehydration,
        }

        public static AzureCPLogging Local
        {
            get
            {
                var LogSvc = SPDiagnosticsServiceBase.GetLocal<AzureCPLogging>();
                // if the Logging Service is registered, just return it.
                if (LogSvc != null)
                    return LogSvc;

                AzureCPLogging svc = null;
                SPSecurity.RunWithElevatedPrivileges(delegate()
                {
                    // otherwise instantiate and register the new instance, which requires farm administrator privileges
                    svc = new AzureCPLogging();
                    //svc.Update();
                });
                return svc;
            }
        }

        public AzureCPLogging() : base(DiagnosticsAreaName, SPFarm.Local) { }
        public AzureCPLogging(string name, SPFarm farm) : base(name, farm) { }

        protected override IEnumerable<SPDiagnosticsArea> ProvideAreas() { yield return Area; }
        public override string DisplayName { get { return DiagnosticsAreaName; } }

        public SPDiagnosticsCategory this[Categories id]
        {
            get { return Areas[DiagnosticsAreaName].Categories[id.ToString()]; }
        }

        public static void WriteTrace(Categories Category, TraceSeverity Severity, string message)
        {
            Local.WriteTrace(1337, Local.GetCategory(Category), Severity, message);
        }

        public static void WriteEvent(Categories Category, EventSeverity Severity, string message)
        {
            Local.WriteEvent(1337, Local.GetCategory(Category), Severity, message);
        }

        public static string FormatException(Exception ex)
        {
            return String.Format("{0}  Stack trace: {1}", ex.Message, ex.StackTrace);
        }

        public static void Unregister()
        {
            SPFarm.Local.Services
                        .OfType<AzureCPLogging>()
                        .ToList()
                        .ForEach(s =>
                        {
                            s.Delete();
                            s.Unprovision();
                            s.Uncache();
                        });
        }

        #region Init categories in area
        private static SPDiagnosticsArea Area
        {
            get
            {
                return new SPDiagnosticsArea(
                    DiagnosticsAreaName,
                    new List<SPDiagnosticsCategory>()
                    {
                        CreateCategory(Categories.Claims_Picking),
                        CreateCategory(Categories.Configuration),
                        CreateCategory(Categories.Lookup),
                        CreateCategory(Categories.Core),
                        CreateCategory(Categories.Claims_Augmentation),
                        CreateCategory(Categories.Rehydration),
                    }
                );
            }
        }

        private static SPDiagnosticsCategory CreateCategory(Categories category)
        {
            return new SPDiagnosticsCategory(
                        GetCategoryName(category),
                        GetCategoryDefaultTraceSeverity(category),
                        GetCategoryDefaultEventSeverity(category)
                    );
        }

        private SPDiagnosticsCategory GetCategory(Categories cat)
        {
            return base.Areas[DiagnosticsAreaName].Categories[GetCategoryName(cat)];
        }

        private static string GetCategoryName(Categories cat)
        {
            // Get the type
            Type type = cat.GetType();
            // Get fieldinfo for this type
            System.Reflection.FieldInfo fieldInfo = type.GetField(cat.ToString());
            // Get the stringvalue attributes
            CategoryNameAttribute[] attribs = fieldInfo.GetCustomAttributes(typeof(CategoryNameAttribute), false) as CategoryNameAttribute[];
            // Return the first if there was a match.
            return attribs.Length > 0 ? attribs[0].Name : null;
        }

        private static TraceSeverity GetCategoryDefaultTraceSeverity(Categories cat)
        {
            // Get the type
            Type type = cat.GetType();
            // Get fieldinfo for this type
            System.Reflection.FieldInfo fieldInfo = type.GetField(cat.ToString());
            // Get the stringvalue attributes
            DefaultTraceSeverityAttribute[] attribs = fieldInfo.GetCustomAttributes(typeof(DefaultTraceSeverityAttribute), false) as DefaultTraceSeverityAttribute[];
            // Return the first if there was a match.
            return attribs.Length > 0 ? attribs[0].Severity : TraceSeverity.Unexpected;
        }

        private static EventSeverity GetCategoryDefaultEventSeverity(Categories cat)
        {
            // Get the type
            Type type = cat.GetType();
            // Get fieldinfo for this type
            System.Reflection.FieldInfo fieldInfo = type.GetField(cat.ToString());
            // Get the stringvalue attributes
            DefaultEventSeverityAttribute[] attribs = fieldInfo.GetCustomAttributes(typeof(DefaultEventSeverityAttribute), false) as DefaultEventSeverityAttribute[];
            // Return the first if there was a match.
            return attribs.Length > 0 ? attribs[0].Severity : EventSeverity.Error;
        }
        #endregion

        #region Attributes
        private class CategoryNameAttribute : Attribute
        {
            public string Name { get; private set; }
            public CategoryNameAttribute(string Name) { this.Name = Name; }
        }

        private class DefaultTraceSeverityAttribute : Attribute
        {
            public TraceSeverity Severity { get; private set; }
            public DefaultTraceSeverityAttribute(TraceSeverity severity) { this.Severity = severity; }
        }

        private class DefaultEventSeverityAttribute : Attribute
        {
            public EventSeverity Severity { get; private set; }
            public DefaultEventSeverityAttribute(EventSeverity severity) { this.Severity = severity; }
        }
        #endregion
    }
}
