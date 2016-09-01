using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PluginContracts;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;

namespace mainUI
{
    public static class PluginLoader
    {
        public static ICollection<IPlugin> LoadPlugins(string path)
        {
            string[] dllFileNames = null;

            if (Directory.Exists(path))
            {
                dllFileNames = Directory.GetFiles(path, "*.dll");

                Console.WriteLine("List of assemblies:");

                ICollection<Assembly> assemblies = new List<Assembly>(dllFileNames.Length);
                foreach (string dllFile in dllFileNames)
                {
                    try
                    {
                        bool contains = Regex.IsMatch(dllFile, "Test", RegexOptions.IgnoreCase);
                        if (contains)
                        {
                            AssemblyName an = AssemblyName.GetAssemblyName(dllFile);
                            Assembly assembly = Assembly.Load(an);
                            assemblies.Add(assembly);
                            Console.WriteLine(dllFile);
                            Console.WriteLine(assembly.FullName);
                        }

                    }
                    catch
                    {

                    }
                }

                Console.WriteLine("List of assemblies types:");
                Type pluginType = typeof(IPlugin);
                ICollection<Type> pluginTypes = new List<Type>();
                foreach (Assembly assembly in assemblies)
                {
                    Console.WriteLine(assembly.FullName);
                    Console.WriteLine(assembly.ImageRuntimeVersion);

                    Regex regex = new Regex(@"test|Test|TEST");
                    Match match = regex.Match(assembly.FullName);
                    if (match.Success)
                    {
                        Type[] types = assembly.GetTypes();
                        foreach (Type type in types)
                        {
                            if (type.IsInterface || type.IsAbstract)
                            {
                                continue;
                            }
                            else
                            {
                                if (type.GetInterface(pluginType.FullName) != null)
                                {
                                    pluginTypes.Add(type);
                                }
                            }
                        }
                    }
                }

                ICollection<IPlugin> plugins = new List<IPlugin>(pluginTypes.Count);
                foreach (Type type in pluginTypes)
                {
                    IPlugin plugin = (IPlugin)Activator.CreateInstance(type);
                    plugins.Add(plugin);
                    Console.WriteLine(plugin.Name);
                }

                return plugins;
            }

            return null;
        }
    }
}
