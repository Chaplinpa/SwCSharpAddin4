using System;

public class OpenInExplorer
{
	public OpenInExplorer()
	{
        ModelDoc2 swModel = iSwApp.ActiveDoc;
        string swModelPath = swModel.GetPathName();
        swModelPath = "/select, " + swModelPath;

        System.Diagnostics.Process Process = new System.Diagnostics.Process();
        Process.StartInfo.UseShellExecute = true;
        Process.StartInfo.FileName = @"explorer";
        Process.StartInfo.Arguments = swModelPath;
        Process.Start();
    }
}
