using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using Xbim.Ifc.ViewModels;
using Xbim.Ifc2x3.Interfaces;
using Xbim.ModelGeometry.Scene;
using Xbim.Common;
using Xbim.Ifc;
using Xbim.Ifc4.Interfaces;
using IIfcProject = Xbim.Ifc4.Interfaces.IIfcProject;
using Xbim.Common.Geometry;
using System.Text;
using System.Runtime.Serialization.Json;
using Xbim.Ifc2x3.ProductExtension;
using System.Diagnostics;
using Newtonsoft.Json;

namespace XbimXplorer.Dialogs
{
    /// <summary>
    /// 
    /// </summary>
    public partial class ExportWindow
    {
        /// <summary>
        /// Interaction logic for ExportWindow.xaml
        /// </summary>
        public ExportWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="callingWindow"></param>
        public ExportWindow(XplorerMainWindow callingWindow) : this()
        {
            _mainWindow = callingWindow;
            TxtFolderName.Text = Path.Combine(
                new FileInfo(_mainWindow.GetOpenedModelFileName()).DirectoryName,
                "Export"
                );
        }

        private XplorerMainWindow _mainWindow;

        private void DoExport(object sender, RoutedEventArgs e)
        {
            var totExports = (ChkWexbim.IsChecked.HasValue && ChkWexbim.IsChecked.Value ? 1 : 0);
            if (totExports == 0)
                return;

            //创建导出目录
            string directName = TxtFolderName.Text;
            if (!Directory.Exists(directName))
            {
                try
                {
                    Directory.CreateDirectory(directName);
                }
                catch (Exception)
                {
                    MessageBox.Show("Error creating directory. Select a different location.");
                    return;
                }
            }

            Cursor = Cursors.Wait;
            if (ChkWexbim.IsChecked.HasValue && ChkWexbim.IsChecked.Value)
            {
                try
                {
                    //导出模型到Wexbim
                    var wexbimFileName = GetExportName("wexbim");
                    using (var wexBimFile = new FileStream(wexbimFileName, FileMode.Create))
                    {
                        using (var binaryWriter = new BinaryWriter(wexBimFile))
                        {
                            try
                            {
                                string []arr = this.TxtTranslation.Text.Split(',');
                                IVector3D translation = new XbimVector3D(double.Parse(arr[0]), double.Parse(arr[1]), double.Parse(arr[2]));
                                _mainWindow.Model.SaveAsWexBim(binaryWriter, null, translation);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }
                            finally
                            {
                                binaryWriter.Flush();
                                wexBimFile.Close();
                            }
                        }
                    }

                    //导出模型结构
                    var propertiesFileName = GetExportName("properties");
                    var project = _mainWindow.Model.Instances.OfType<IIfcProject>().FirstOrDefault();
                    List<ModelInfo> lstModel = new List<ModelInfo>();
                    foreach (var item in project.SpatialStructuralElements)
                    {
                        var sv = new SpatialViewModel(item, null);
                        ModelInfo model = ParseModelInfo(sv);
                        lstModel.Add(model);
                    }
                    string strProperties = JsonConvert.SerializeObject(lstModel);
                    File.WriteAllText(propertiesFileName, strProperties);

                    Process.Start(this.TxtFolderName.Text);
                    MessageBox.Show("Ifc Export Wexbim Sucessed!");
                }
                catch (Exception ce)
                {
                    if (CancelAfterNotification("Error exporting Wexbim file.", ce, totExports))
                    {
                        Cursor = Cursors.Arrow;
                        return;
                    }
                }
                totExports--;
            }
            Cursor = Cursors.Arrow;
            Close();
        }

        //选择输出目录
        private void SelectOutputPath(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog folderDlg = new System.Windows.Forms.FolderBrowserDialog();
            folderDlg.SelectedPath = this.TxtFolderName.Text;

            if (folderDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                TxtFolderName.Text = folderDlg.SelectedPath;
            }
        }

        /// <summary>
		/// Recursion  parse ifcmodel info to property model
		/// </summary>
		/// <param name="model"></param>
		/// <returns></returns>
        private ModelInfo ParseModelInfo(IXbimViewModel model)
        {
            if (model == null)
                return null;

            IPersistEntity site = model.Entity;

            ModelInfo info = new ModelInfo();
            info.Name = model.Name;
            info.GlobalId = (site as Xbim.Ifc2x3.Kernel.IfcRoot).GlobalId.ToString();
            info.ProductId = site.EntityLabel.ToString();

            if (model.Children.Count() > 0)
            {
                foreach (IXbimViewModel cNode in model.Children)
                {
                    ModelInfo child = ParseModelInfo(cNode as IXbimViewModel);
                    if (child != null)
                    {
                        info.Children.Add(child);
                    }
                }
            }
            return info;
        }

        private string GetExportName(string extension, int progressive = 0)
        {
            var basefile = new FileInfo(_mainWindow.GetOpenedModelFileName());
            var wexbimFileName = Path.Combine(TxtFolderName.Text, basefile.Name);
            if (progressive != 0)
                extension = progressive + "." + extension;
            wexbimFileName = Path.ChangeExtension(wexbimFileName, extension);
            return wexbimFileName;
        }

        private bool CancelAfterNotification(string errorZoneMessage, Exception ce, int totExports)
        {
            var tasksLeft = totExports - 1;
            var message = errorZoneMessage + "\r\n" + ce.Message + "\r\n";

            if (tasksLeft > 0)
            {
                message += "\r\n" +
                           string.Format(
                               "Do you wish to continue exporting other formats?", tasksLeft
                               );
                var ret = MessageBox.Show(message, "Error", MessageBoxButton.YesNoCancel, MessageBoxImage.Error);
                return ret != MessageBoxResult.Yes;
            }
            else
            {
                var ret = MessageBox.Show(message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return ret != MessageBoxResult.Yes;
            }
        }
    }

    public class ModelInfo
    {
        public ModelInfo()
        {
            this.Children = new List<ModelInfo>();
        }

        public ModelInfo(string globalId, string name, string productId) : this()
        {
            this.GlobalId = globalId;
            this.Name = name;
            this.ProductId = productId;
        }

        public string GlobalId { get; set; }
        public string Name { get; set; }
        public string ProductId { get; set; }
        public List<ModelInfo> Children { get; set; }
    }
}
