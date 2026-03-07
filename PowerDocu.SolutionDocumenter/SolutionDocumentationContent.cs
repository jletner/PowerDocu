using System;
using System.Collections.Generic;
using System.Linq;
using PowerDocu.Common;

namespace PowerDocu.SolutionDocumenter
{
    public class SolutionDocumentationContent
    {
        public List<FlowEntity> flows = new List<FlowEntity>();
        public List<AppEntity> apps = new List<AppEntity>();
        public List<AppModuleEntity> appModules = new List<AppModuleEntity>();
        public SolutionEntity solution;
        public string folderPath,
            filename;

        public SolutionDocumentationContent(
            SolutionEntity solution,
            List<AppEntity> apps,
            List<FlowEntity> flows,
            List<AppModuleEntity> appModules,
            string path
        )
        {
            this.solution = solution;
            this.apps = apps;
            this.flows = flows;
            this.appModules = appModules ?? new List<AppModuleEntity>();
            filename = CharsetHelper.GetSafeName(solution.UniqueName);
            folderPath = path;
        }

        public string GetDisplayNameForComponent(SolutionComponent component)
        {
            if (component.Type == "Workflow")
            {
                //todo this seems to be always null at the moment, need to investigate
                FlowEntity flow = flows.Where(f => f.Name.Contains(component.ID, StringComparison.OrdinalIgnoreCase))?.FirstOrDefault();
                if (flow != null)
                {
                    return flow.Name + " (" + flow.trigger.Name + ": " + flow.trigger.Type + ")";
                }
            }
            if (component.Type == "Model-Driven App")
            {
                AppModuleEntity appModule = appModules.Where(a => a.UniqueName != null && a.UniqueName.Equals(component.SchemaName, StringComparison.OrdinalIgnoreCase))?.FirstOrDefault();
                if (appModule != null)
                {
                    return appModule.GetDisplayName();
                }
            }
            return solution.GetDisplayNameForComponent(component);
        }
    }
}
