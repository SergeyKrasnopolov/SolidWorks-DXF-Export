using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;

[assembly: ComVisible(false)]

namespace SolidWorksDxfExport
{
    [ComVisible(true)]
    [Guid("92D5D5E8-7B21-4A54-B322-3AFC73B7E960")]
    [ProgId("SolidWorksDxfExport.FrontView")]
    [ClassInterface(ClassInterfaceType.None)]
    public sealed class FrontViewAddin : ISwAddin
    {
        public bool ConnectToSW(object ThisSW, int Cookie)
        {
            try
            {
                ISldWorks app = (ISldWorks)ThisSW;
                IModelDoc2 model = app.ActiveDoc as IModelDoc2;
                if (model == null)
                {
                    MessageBox.Show("Откройте деталь или сборку, затем выключите и снова включите надстройку.",
                        "DXF Export — вид спереди");
                    return true;
                }

                int type = model.GetType();
                if (type != (int)swDocumentTypes_e.swDocPART &&
                    type != (int)swDocumentTypes_e.swDocASSEMBLY)
                {
                    MessageBox.Show("Вид спереди доступен для детали или сборки. Активный документ — не 3D-модель.",
                        "DXF Export — вид спереди");
                    return true;
                }

                // Use the standard view ID instead of a localized view name.
                model.ShowNamedView2("", (int)swStandardViews_e.swFront);
                model.GraphicsRedraw2();
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Не удалось включить вид спереди:\n" + ex.Message,
                    "DXF Export — ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        public bool DisconnectFromSW()
        {
            // No event handlers or retained COM references in this first version.
            return true;
        }

        [ComRegisterFunction]
        public static void Register(Type type)
        {
            using (RegistryKey key = Registry.LocalMachine.CreateSubKey(
                @"SOFTWARE\SolidWorks\Addins\{" + type.GUID.ToString() + "}"))
            {
                key.SetValue(null, 0, RegistryValueKind.DWord);
                key.SetValue("Title", "DXF Export - Front View");
                key.SetValue("Description", "Set the active part or assembly to Front view when enabled.");
            }
        }

        [ComUnregisterFunction]
        public static void Unregister(Type type)
        {
            Registry.LocalMachine.DeleteSubKey(
                @"SOFTWARE\SolidWorks\Addins\{" + type.GUID.ToString() + "}", false);
            Registry.CurrentUser.DeleteSubKey(
                @"Software\SolidWorks\AddInsStartup\{" + type.GUID.ToString() + "}", false);
        }
    }
}
