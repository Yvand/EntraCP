using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using System;
using System.Collections.Generic;
using System.Diagnostics;
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

        public enum TraceCategory
        {
            [CategoryName("Core"),
             DefaultTraceSeverity(TraceSeverity.Medium),
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
            [CategoryName("Rehydration"),
             DefaultTraceSeverity(TraceSeverity.Medium),
             DefaultEventSeverity(EventSeverity.Error)]
            Rehydration,
            [CategoryName("Augmentation"),
             DefaultTraceSeverity(TraceSeverity.Medium),
             DefaultEventSeverity(EventSeverity.Error)]
            Augmentation,
            [CategoryName("Debug"),
             DefaultTraceSeverity(TraceSeverity.Medium),
             DefaultEventSeverity(EventSeverity.Error)]
            Debug,
            [CategoryName("Custom"),
             DefaultTraceSeverity(TraceSeverity.Medium),
             DefaultEventSeverity(EventSeverity.Error)]
            Custom,
        }

        public static void Log(string message, TraceSeverity traceSeverity, EventSeverity eventSeverity, TraceCategory category)
        {
            try
            {
                WriteTrace(category, traceSeverity, message);
                //LdapcpLoggingService.WriteEvent(LdapcpLoggingService.Categories.LDAPCP, eventSeverity, message);
            }
            catch
            {   // Don't want to do anything if logging goes wrong, just ignore and continue
            }
        }

        public static void LogException(string ProviderInternalName, string faultyAction, TraceCategory category, Exception ex)
        {
            try
            {
                string message = "[{0}] Unexpected error {1}: {2}: {3}, Callstack: {4}";
                if (ex is AggregateException)
                {
                    var aggEx = ex as AggregateException;
                    foreach (var innerEx in aggEx.InnerExceptions)
                    {
                        if (innerEx.InnerException != null)
                            message = String.Format(message, ProviderInternalName, faultyAction, innerEx.InnerException.GetType().FullName, innerEx.InnerException.Message, innerEx.InnerException.StackTrace);
                        else
                            message = String.Format(message, ProviderInternalName, faultyAction, innerEx.GetType().FullName, innerEx.Message, innerEx.StackTrace);
                        WriteTrace(category, TraceSeverity.Unexpected, message);
                    }
                }
                else
                {
                    if (ex.InnerException != null)
                        message = String.Format(message, ProviderInternalName, faultyAction, ex.InnerException.GetType().FullName, ex.InnerException.Message, ex.InnerException.StackTrace);
                    else
                        message = String.Format(message, ProviderInternalName, faultyAction, ex.GetType().FullName, ex.Message, ex.StackTrace);
                    WriteTrace(category, TraceSeverity.Unexpected, message);
                }
            }
            catch
            {   // Don't want to do anything if logging goes wrong, just ignore and continue
            }
        }

        /// <summary>
        /// Record message (in VerboseEx) only if assembly is compiled in debug mode
        /// </summary>
        /// <param name="message"></param>
        public static void LogDebug(string message)
        {
            try
            {
                TraceSeverity severity;
#if DEBUG
                severity = TraceSeverity.VerboseEx;
                WriteTrace(TraceCategory.Debug, severity, message);
                Debug.WriteLine(message);
#else
                severity = TraceSeverity.VerboseEx;
#endif

            }
            catch
            {   // Don't want to do anything if logging goes wrong, just ignore and continue
            }
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
                SPSecurity.RunWithElevatedPrivileges(delegate ()
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

        public SPDiagnosticsCategory this[TraceCategory id]
        {
            get { return Areas[DiagnosticsAreaName].Categories[id.ToString()]; }
        }

        public static void WriteTrace(TraceCategory Category, TraceSeverity Severity, string message)
        {
            Local.WriteTrace(1337, Local.GetCategory(Category), Severity, message);
        }

        public static void WriteEvent(TraceCategory Category, EventSeverity Severity, string message)
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
                        CreateCategory(TraceCategory.Claims_Picking),
                        CreateCategory(TraceCategory.Configuration),
                        CreateCategory(TraceCategory.Lookup),
                        CreateCategory(TraceCategory.Core),
                        CreateCategory(TraceCategory.Augmentation),
                        CreateCategory(TraceCategory.Rehydration),
                        CreateCategory(TraceCategory.Debug),
                        CreateCategory(TraceCategory.Custom),
                    }
                );
            }
        }

        private static SPDiagnosticsCategory CreateCategory(TraceCategory category)
        {
            return new SPDiagnosticsCategory(
                        GetCategoryName(category),
                        GetCategoryDefaultTraceSeverity(category),
                        GetCategoryDefaultEventSeverity(category)
                    );
        }

        private SPDiagnosticsCategory GetCategory(TraceCategory cat)
        {
            return base.Areas[DiagnosticsAreaName].Categories[GetCategoryName(cat)];
        }

        private static string GetCategoryName(TraceCategory cat)
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

        private static TraceSeverity GetCategoryDefaultTraceSeverity(TraceCategory cat)
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

        private static EventSeverity GetCategoryDefaultEventSeverity(TraceCategory cat)
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
