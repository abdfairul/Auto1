using System.ComponentModel;
using System.Windows.Forms; // for Ribbon

namespace PluginContracts
{
    public delegate void BeforeExecute();
    public delegate void AfterExecute();
    public delegate void BeforeExecuteButAfterSelection(object selection, object datagridview, object owner, BackgroundWorker backgroundWorker);
    public delegate void CellClickEvent(object sender, DataGridViewCellEventArgs e);
    public delegate void ProgressBarChangedEvent(object sender, ProgressChangedEventArgs e);
    public interface IPlugin
	{
		string Name { get; }
        RibbonTab EquipmentSetting { get; }
        DataGridViewRow ToExecute { get; set; }
        PictureBox Picture { get; set; }
        bool Busy { get; }   
        BeforeExecute BeforeExecute { get; }
        AfterExecute AfterExecute { get; }
        BeforeExecuteButAfterSelection BeforeExecuteButAfterSelection { get; }
        ProgressBarChangedEvent ProgressBarChangedEvent { get; }
        bool DoExecute { get; }
        CellClickEvent ClickEvent { get; }
    }
}
