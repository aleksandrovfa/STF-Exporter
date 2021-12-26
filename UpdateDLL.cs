using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Эталон_Проект
{
    class UpdateDLL
    {
        public Assembly assembly { get; set; }
        public object ObjectInstance { get; set; }
        public UpdateDLL(string location, string dll_name, string type_name, Object[] objects)//ExternalCommandData commandData)//ExternalCommandData commandData, string dll_name, string type_name)
        {
            string path_active_addin = location + "\\" + dll_name + ".dll",//CopyElements.dll//location + @"\Эталон Проект\"
                path_name_dll = "";
               // path_file_load_addin = @"O:\Etalon Project\KPO\Revit\!_БИБЛИОТЕКА РЕСУРСОВ\07_ПЛАГИНЫ\ЭталонПроект\DLL\" + dll_name + ".dll";//CopyElements.dll";
            if (!double.TryParse(FileVersionInfo.GetVersionInfo(path_active_addin).ProductVersion.Replace(".", ""), out double resold))
            {

            }
            foreach (string file_name in Directory.GetFiles(location))// + @"\Эталон Проект\"))
            {
                if (file_name.Contains(dll_name))//"CopyElements"))
                {
                    double.TryParse(FileVersionInfo.GetVersionInfo(file_name).ProductVersion.Replace(".", ""), out double resnew);
                    //FileVersionInfo VersionNew = FileVersionInfo.GetVersionInfo(path_file_load_addin);
                    if (resnew >= resold)//File.GetCreationTime(path_file_load_addin).ToString() == File.GetCreationTime(file_name).ToString())
                    {
                        resold = resnew;
                        path_name_dll = file_name;
                    }
                }
            }
            //if (path_name_dll == "")
            //{
            //    try
            //    {
            //        File.Copy(path_file_load_addin, path_active_addin);
            //        path_name_dll = path_active_addin;
            //    }
            //    catch
            //    {
            //        path_active_addin = path_active_addin.Replace(dll_name + ".dll", dll_name + "1.dll");//"CopyElements.dll", "CopyElements1.dll");
            //        for (int i = 0; ; i++)
            //        {
            //            path_active_addin = path_active_addin.Replace(((i - 1).ToString() + ".dll"), (i.ToString() + ".dll"));
            //            try
            //            {
            //                File.Copy(path_file_load_addin, path_active_addin);
            //                path_name_dll = path_active_addin;
            //                break;
            //            }
            //            catch { }
            //        }
            //    }
            //}
            assembly = Assembly.LoadFile(path_name_dll);
            ObjectInstance = assembly.CreateInstance(type_name);//"CopyElements.CopyGroupForPipe"
            if (objects != null)//(commandData != null)
            {
                Type t = assembly.GetType(type_name);//"CopyElements.CopyGroupForPipe"
                //Object[] objects = new Object[1];
                //objects[0] = commandData;
                MethodInfo methodinfo = t.GetMethod("Execute");
                methodinfo.Invoke(ObjectInstance, objects);
            }
            //else
            //{
            //    assembly = Assembly.LoadFile(path_name_dll);
            //}
        }
        public void TryAgain(Object[] objects)
        {
            Type t = ObjectInstance.GetType();
            MethodInfo methodinfo = t.GetMethod("Execute");
            methodinfo.Invoke(ObjectInstance, objects);
        }
    }
    class Bardage
    {
        public bool UpdateNeed = false;
        public string path_name = "";
        string path_addin = @"O:\Etalon Project\KPO\Revit\!_БИБЛИОТЕКА РЕСУРСОВ\07_ПЛАГИНЫ\ЭталонПроект\DLL\__", path_location_user = "";
        public Bardage(UIControlledApplication a)
        {
            path_location_user = a.ControlledApplication.AllUsersAddinsLocation + @"\Эталон Проект";
            foreach (string st in Directory.GetFiles(a.ControlledApplication.AllUsersAddinsLocation))
            {
                if (st.Contains("Эталон Проект"))
                {
                    path_name = Path.GetFileName(st).Replace(".addin", "");
                }
            }
        }
        public Bardage(string location)//UIControlledApplication a)
        {
            path_location_user = location + @"\Эталон Проект";//a.ControlledApplication.AllUsersAddinsLocation + @"\Эталон Проект";
            foreach (string st in Directory.GetFiles(location))
            {
                if (st.Contains("Эталон Проект"))
                {
                    path_name = Path.GetFileName(st).Replace(".addin", "");
                }
            }
        }
        public bool ClearBardage(string dll_name, string panel_name)
        {
            try
            {
                string origin_name = path_location_user + "\\" + panel_name + "\\" + dll_name + ".dll";
                List<string> files_name = Directory.GetFiles(path_location_user + "\\" + panel_name).ToList();
                string product_name = FileVersionInfo.GetVersionInfo(path_location_user + "\\" + panel_name + "\\" + dll_name + ".dll").ProductName;
                List<string> files_name_byProduct = files_name.Where(x => FileVersionInfo.GetVersionInfo(x).ProductName == product_name).ToList();
                if (double.TryParse(FileVersionInfo.GetVersionInfo(origin_name).ProductVersion.Replace(".", ""), out double result_old))
                {
                    foreach (string st in files_name_byProduct)
                    {
                        if(st == origin_name)
                        {
                            continue;
                        }
                        if (double.TryParse(FileVersionInfo.GetVersionInfo(st).ProductVersion.Replace(".", ""), out double result_new))
                        {
                            if (result_new > result_old)
                            {
                                //File.Move(st, path_location_user + @"\" + dll_name + ".dll");
                                string fileNameNew = st.Substring(0, st.LastIndexOf(dll_name)) + dll_name + ".dll";
                                if (File.Exists(fileNameNew))
                                {
                                    File.Delete(fileNameNew);
                                }
                                File.Move(st, fileNameNew);
                                result_old = result_new;
                            }
                            File.Delete(st);
                        }
                    }
                    if (double.TryParse(FileVersionInfo.GetVersionInfo(path_addin + path_name + "\\__" + panel_name + "\\" + dll_name + ".dll").ProductVersion.Replace(".", ""), out double result1))
                    {
                        if (result1 > result_old)
                        {
                            //System.Windows.Forms.MessageBox.Show(path_addin + path_name + "\\__" + panel_name + "\\" + dll_name + ".dll");
                            //System.Windows.Forms.MessageBox.Show(path_addin + path_name + "\\__" + panel_name + "\\" + dll_name);
                            UpdateNeed = true;
                        }
                    }
                }
                //System.Windows.Forms.MessageBox.Show(path_addin + path_name + "\\__" + panel_name + "\\" + dll_name + ".dll");
                //if (double.TryParse(FileVersionInfo.GetVersionInfo(path_addin + path_name + "\\__" + panel_name + "\\" + dll_name + ".dll").ProductVersion.Replace(".", ""), out double result1))
                //{
                //    if (result1 > result_old)
                //    {
                //        System.Windows.Forms.MessageBox.Show(path_addin + path_name + "\\__" + panel_name + "\\" + dll_name + ".dll");
                //        //System.Windows.Forms.MessageBox.Show(path_addin + path_name + "\\__" + panel_name + "\\" + dll_name);
                //        UpdateNeed = true;
                //    }
                //}
                //
                //List<string> files_name = Directory.GetFiles(path_location_user + "\\" + panel_name).ToList();
                ////List<string> files_name = Directory.GetFiles(path_location_user + "\\" + panel_name + "\\" ).ToList();
                //var alllist = files_name
                //    .ConvertAll(x => FileVersionInfo.GetVersionInfo(x).OriginalFilename)
                //    .GroupBy(x => x)
                //    .ToList();
                //if (!alllist.ConvertAll(x => x.Key).Contains(dll_name + ".dll"))
                //{
                //    return false;
                //}
                ////List<string> list =
                //    string origin_name =
                //    alllist
                //    .Where(x => x.Key == FileVersionInfo.GetVersionInfo(path_location_user + "\\" + panel_name + "\\" + dll_name + ".dll").OriginalFilename)//&& x.Key != null x.Count() > 1 && 
                //    .ToList()
                //    .ConvertAll(x => x.Key).First();
                //double NumberVersion = 0;
                ////foreach (string origin_name in list)
                ////{
                //    foreach (string st in files_name.Where(x => x.Contains(origin_name.Replace(".dll", ""))).ToList())
                //    {
                //        if (double.TryParse(FileVersionInfo.GetVersionInfo(st).ProductVersion.Replace(".", ""), out double result))
                //        {
                //            if (result > NumberVersion)// && st.Contains(dll_name + ".dll"))
                //            {
                //                NumberVersion = result;
                //            if(!st.Contains(dll_name + ".dll"))
                //            {
                //                File.Move(st, path_location_user + @"\" + dll_name + ".dll");
                //            }
                //        }
                //            else if (!st.Contains(dll_name + ".dll"))//(!st.Contains(origin_name))
                //            {
                //                File.Delete(st);
                //            }
                //        }
                //    }
                //    if (double.TryParse(FileVersionInfo.GetVersionInfo(path_name + "\\__" + panel_name + "\\__" + origin_name + ".dll").ProductVersion.Replace(".", ""), out double result1))
                //    {
                //        if(result1 > NumberVersion)
                //        {
                //            UpdateNeed = true;
                //        }
                //    }
                ////}
            }
            catch
            {
                return false;
            }
            return true;
        }
        public bool ClearBardage(string dll_name, string panel_name, DLL dll)
        {
            try
            {
                string origin_name = path_location_user + "\\" + panel_name + "\\" + dll_name + ".dll";
                List<string> files_name = Directory.GetFiles(path_location_user + "\\" + panel_name).ToList();
                string product_name = FileVersionInfo.GetVersionInfo(path_location_user + "\\" + panel_name + "\\" + dll_name + ".dll").ProductName;
                List<string> files_name_byProduct = files_name.Where(x => FileVersionInfo.GetVersionInfo(x).ProductName == product_name).ToList();
                if (double.TryParse(FileVersionInfo.GetVersionInfo(origin_name).ProductVersion.Replace(".", ""), out double result_old))
                {
                    foreach (string st in files_name_byProduct)
                    {
                        if (st == origin_name)
                        {
                            continue;
                        }
                        FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(st);
                        if (double.TryParse(fvi.ProductVersion.Replace(".", ""), out double result_new))
                        {
                            if (result_new > result_old)
                            {
                                //File.Move(st, path_location_user + @"\" + dll_name + ".dll");
                                result_old = result_new;
                                dll.version_old = result_new;
                                dll.Version_Old = fvi.ProductVersion;
                                //dll.version_new

                            }
                            //File.Delete(st);
                        }
                    }
                }
                //if (double.TryParse(FileVersionInfo.GetVersionInfo(path_addin + path_name + "\\__" + panel_name + "\\" + dll_name + ".dll").ProductVersion.Replace(".", ""), out double result1))
                //{
                //    if (result1 > result_old)
                //    {
                //        dll.version_new = result;
                //        //UpdateNeed = true;
                //    }
                //}
                //List<string> files_name = Directory.GetFiles(path_location_user + "\\" + panel_name).ToList();
                //var alllist = files_name
                //    .ConvertAll(x => FileVersionInfo.GetVersionInfo(x).OriginalFilename)
                //    .GroupBy(x => x)
                //    .ToList();
                //if (!alllist.ConvertAll(x => x.Key).Contains(dll_name + ".dll"))
                //{
                //    return false;
                //}
                //List<string> list =
                //    alllist
                //    .Where(x => x.Count() > 1 && x.Key == FileVersionInfo.GetVersionInfo(path_location_user + "\\" + panel_name + "\\" + dll_name + ".dll").OriginalFilename)
                //    .ToList()
                //    .ConvertAll(x => x.Key);
                //foreach (string origin_name in list)
                //{
                //    foreach (string st in files_name.Where(x => x.Contains(origin_name.Replace(".dll", ""))).ToList())
                //    {
                //        if (double.TryParse(FileVersionInfo.GetVersionInfo(st).ProductVersion.Replace(".", ""), out double result))
                //        {
                //            if (result > dll.version_new && st.Contains(dll_name + ".dll"))
                //            {
                //                dll.version_new = result;
                //                File.Move(st, path_location_user + @"\" + dll_name + ".dll");
                //            }
                //        }
                //    }
                //}
            }
            catch
            {
                return false;
            }
            return true;
        }
    }
}
